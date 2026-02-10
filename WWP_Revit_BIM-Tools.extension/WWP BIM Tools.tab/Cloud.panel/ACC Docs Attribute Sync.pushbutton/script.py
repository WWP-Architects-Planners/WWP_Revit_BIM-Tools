#! python3
import json
import os
import re
import traceback
from pyrevit import UI
from pyrevit import script as pyrevit_script
from System import DateTime, Guid, Type, Activator, Array, Object, Uri, Convert, TimeSpan, Action
from System.Reflection import BindingFlags
from System.Runtime.InteropServices import Marshal
from System.Net import ServicePointManager, SecurityProtocolType, HttpListener, WebRequest, WebException
from System.Text import Encoding
from System.IO import StreamReader, FileStream, FileMode, FileAccess
from System.Diagnostics import Process, ProcessStartInfo
from System.Threading import Thread, ThreadStart
from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
from System.Windows.Controls import TreeViewItem
from System.Windows.Input import MouseButtonEventHandler
from System.Windows import RoutedEventHandler
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Markup import XamlReader
from Microsoft.Win32 import OpenFileDialog

TITLE = "ACC Docs"
REDIRECT_URI = "http://127.0.0.1:8765/callback/"
OAUTH_AUTHORIZE_URL = "https://developer.api.autodesk.com/authentication/v2/authorize"
OAUTH_TOKEN_URL = "https://developer.api.autodesk.com/authentication/v2/token"
DEFAULT_SCOPES = "data:read data:write account:read"

ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

ENV_CLIENT_ID_KEY = "CLIENT_ID"
ENV_CLIENT_SECRET_KEY = "CLIENT_SECRET"


def _load_env_file(path):
    data = {}
    if not path or not os.path.exists(path):
        return data
    try:
        with open(path, "r") as fh:
            for line in fh:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip().strip('"').strip("'")
                data[key] = value
    except Exception:
        return data
    return data


def _get_env_credentials():
    creds = {}
    # prefer environment variables
    if os.environ.get(ENV_CLIENT_ID_KEY):
        creds[ENV_CLIENT_ID_KEY] = os.environ.get(ENV_CLIENT_ID_KEY)
    if os.environ.get(ENV_CLIENT_SECRET_KEY):
        creds[ENV_CLIENT_SECRET_KEY] = os.environ.get(ENV_CLIENT_SECRET_KEY)

    # fallback to .env file next to script or extension root
    script_dir = os.path.dirname(__file__)
    env_paths = [
        os.path.join(script_dir, ".env"),
        os.path.join(os.path.dirname(script_dir), ".env"),
    ]
    for env_path in env_paths:
        data = _load_env_file(env_path)
        for key in (ENV_CLIENT_ID_KEY, ENV_CLIENT_SECRET_KEY):
            if key in data and key not in creds:
                creds[key] = data[key]
    return creds


class ComInterop(object):
    FLAGS = BindingFlags.Public | BindingFlags.Instance | BindingFlags.OptionalParamBinding

    @staticmethod
    def get(target, name, *args):
        if target is None:
            return None
        return target.GetType().InvokeMember(
            name,
            ComInterop.FLAGS | BindingFlags.GetProperty,
            None,
            target,
            args,
        )

    @staticmethod
    def set(target, name, value):
        if target is None:
            return
        target.GetType().InvokeMember(
            name,
            ComInterop.FLAGS | BindingFlags.SetProperty,
            None,
            target,
            [value],
        )

    @staticmethod
    def call(target, name, *args):
        if target is None:
            return None
        return target.GetType().InvokeMember(
            name,
            ComInterop.FLAGS | BindingFlags.InvokeMethod,
            None,
            target,
            args,
        )

    @staticmethod
    def release(obj):
        try:
            if obj is not None and Marshal.IsComObject(obj):
                Marshal.FinalReleaseComObject(obj)
        except Exception:
            pass


def coerce_to_2d(values):
    if isinstance(values, Array):
        return values
    arr = Array.CreateInstance(Object, 1, 1)
    arr[0, 0] = values
    return arr


class ExcelRow(object):
    def __init__(self, file_name="", description=""):
        self.FileName = file_name
        self.Description = description


class ExcelReader(object):
    LastDiagnostics = None
    MaxLogRows = 20

    @staticmethod
    def _normalize_name(value):
        if value is None:
            return ""
        try:
            text = str(value)
        except Exception:
            return ""
        try:
            text = text.replace(u"\u00A0", " ")
            text = text.replace(u"\u200e", "")
            text = text.replace(u"\u200f", "")
            text = text.replace(u"\u202a", "").replace(u"\u202b", "").replace(u"\u202c", "")
            text = text.replace(u"\u2013", "-").replace(u"\u2014", "-").replace(u"\u2212", "-")
        except Exception:
            pass
        try:
            text = re.sub(r"\s+", " ", text)
        except Exception:
            pass
        return text.strip().lower()

    @staticmethod
    def _normalize_base(value):
        name = ExcelReader._normalize_name(value)
        if not name:
            return ""
        try:
            base, _ext = os.path.splitext(name)
        except Exception:
            return name
        return base if base else name

    @staticmethod
    def _add_result(result, key, row):
        if key and key not in result:
            result[key] = row

    @staticmethod
    def read_excel(path):
        result = {}
        diag = {
            "row_start": None,
            "row_end": None,
            "col_start": None,
            "col_end": None,
            "headers": [],
            "description_index": None,
            "rows_read": 0,
            "samples": [],
            "rows": [],
            "sheet": None,
            "sheet_rows": 0,
            "sheet_desc_rows": 0,
        }
        excel = None
        workbooks = None
        workbook = None
        worksheets = None
        sheet = None
        used_range = None

        excel_type = Type.GetTypeFromProgID("Excel.Application")
        if excel_type is None:
            raise Exception("Excel is not installed.")

        try:
            excel = Activator.CreateInstance(excel_type)
            ComInterop.set(excel, "Visible", False)
            workbooks = ComInterop.get(excel, "Workbooks")
            workbook = ComInterop.call(workbooks, "Open", path)
            worksheets = ComInterop.get(workbook, "Worksheets")
            # pick the sheet with the most data in column A/B
            values = None
            best = {
                "sheet": None,
                "used_range": None,
                "values": None,
                "row_start": None,
                "row_end": None,
                "col_start": None,
                "col_end": None,
                "rows": 0,
                "desc_rows": 0,
            }

            try:
                sheet_count = ComInterop.get(worksheets, "Count")
            except Exception:
                sheet_count = 1

            def safe_get_sheet_text(sht, r, c):
                cell = None
                try:
                    cell = ComInterop.get(sht, "Cells", r, c)
                    return ComInterop.get(cell, "Text")
                except Exception:
                    return None
                finally:
                    try:
                        ComInterop.release(cell)
                    except Exception:
                        pass

            for idx in range(1, int(sheet_count) + 1):
                try:
                    ws = ComInterop.get(worksheets, "Item", idx)
                    rng = ComInterop.get(ws, "UsedRange")
                    vals = coerce_to_2d(ComInterop.get(rng, "Value2"))
                    if vals is None:
                        ComInterop.release(rng)
                        ComInterop.release(ws)
                        continue

                    rs = vals.GetLowerBound(0)
                    re = vals.GetUpperBound(0)
                    cs = vals.GetLowerBound(1)
                    ce = vals.GetUpperBound(1)
                    if re < rs or ce < cs:
                        ComInterop.release(rng)
                        ComInterop.release(ws)
                        continue

                    # Count non-empty in column A and B (row 2+)
                    rows = 0
                    desc_rows = 0
                    for r in range(rs + 1, re + 1):
                        v_a = None
                        v_b = None
                        try:
                            v_a = vals[r, cs]
                        except Exception:
                            v_a = None
                        if v_a is None:
                            v_a = safe_get_sheet_text(ws, r, cs)
                        if v_a is not None and str(v_a).strip():
                            rows += 1
                        if (cs + 1) <= ce:
                            try:
                                v_b = vals[r, cs + 1]
                            except Exception:
                                v_b = None
                            if v_b is None:
                                v_b = safe_get_sheet_text(ws, r, cs + 1)
                            if v_b is not None and str(v_b).strip():
                                desc_rows += 1

                    pick = False
                    if desc_rows > best["desc_rows"]:
                        pick = True
                    elif desc_rows == best["desc_rows"] and rows > best["rows"]:
                        pick = True

                    if pick:
                        # release previous best objects if any
                        try:
                            ComInterop.release(best["used_range"])
                            ComInterop.release(best["sheet"])
                        except Exception:
                            pass
                        best = {
                            "sheet": ws,
                            "used_range": rng,
                            "values": vals,
                            "row_start": rs,
                            "row_end": re,
                            "col_start": cs,
                            "col_end": ce,
                            "rows": rows,
                            "desc_rows": desc_rows,
                        }
                    else:
                        ComInterop.release(rng)
                        ComInterop.release(ws)
                except Exception:
                    try:
                        ComInterop.release(rng)
                        ComInterop.release(ws)
                    except Exception:
                        pass

            sheet = best["sheet"]
            used_range = best["used_range"]
            values = best["values"]

            if values is None:
                ExcelReader.LastDiagnostics = diag
                return result

            row_start = values.GetLowerBound(0)
            row_end = values.GetUpperBound(0)
            col_start = values.GetLowerBound(1)
            col_end = values.GetUpperBound(1)
            diag["row_start"] = row_start
            diag["row_end"] = row_end
            diag["col_start"] = col_start
            diag["col_end"] = col_end
            try:
                diag["sheet"] = ComInterop.get(sheet, "Name")
            except Exception:
                diag["sheet"] = None
            diag["sheet_rows"] = best.get("rows", 0)
            diag["sheet_desc_rows"] = best.get("desc_rows", 0)

            if row_end < row_start or col_end < col_start:
                ExcelReader.LastDiagnostics = diag
                return result

            def safe_get(r, c):
                if r < row_start or r > row_end or c < col_start or c > col_end:
                    return None
                try:
                    return values[r, c]
                except Exception:
                    return None

            def safe_get_text(r, c):
                if sheet is None:
                    return None
                return safe_get_sheet_text(sheet, r, c)

            headers = {}
            header_row = row_start
            for col in range(col_start, col_end + 1):
                header_val = safe_get(header_row, col)
                if header_val is None:
                    continue
                header = ExcelReader._normalize_name(header_val)
                if not header:
                    continue
                headers[header] = (col - col_start) + 1
            try:
                diag["headers"] = [h for h in headers.keys()]
            except Exception:
                diag["headers"] = []

            # Always read descriptions from Column B (row 2+),
            # per user request, regardless of header text.
            description_index = col_start + 1 if (col_start + 1) <= col_end else -1
            diag["description_index"] = description_index

            for row in range(row_start + 1, row_end + 1):
                file_val = safe_get(row, col_start)
                if file_val is None or not str(file_val).strip():
                    file_val = safe_get_text(row, col_start)
                if file_val is None:
                    continue
                file_name = str(file_val).strip()
                if not file_name:
                    continue

                description = ""
                if description_index >= col_start and description_index <= col_end:
                    desc_val = safe_get(row, description_index)
                    if desc_val is None or (isinstance(desc_val, str) and not desc_val.strip()):
                        desc_val = safe_get_text(row, description_index)
                    if desc_val is not None:
                        description = str(desc_val).strip()

                try:
                    if len(diag["samples"]) < 5:
                        diag["samples"].append((file_val, safe_get(row, description_index)))
                except Exception:
                    pass

                row_obj = ExcelRow(file_name, description)
                ExcelReader._add_result(result, ExcelReader._normalize_name(file_name), row_obj)
                ExcelReader._add_result(result, ExcelReader._normalize_base(file_name), row_obj)
                diag["rows_read"] = diag["rows_read"] + 1
                try:
                    if len(diag["rows"]) < ExcelReader.MaxLogRows:
                        diag["rows"].append((file_name, description))
                except Exception:
                    pass

            ExcelReader.LastDiagnostics = diag
            return result
        finally:
            if workbook is not None:
                ComInterop.call(workbook, "Close", False)
            if excel is not None:
                ComInterop.call(excel, "Quit")

            ComInterop.release(used_range)
            ComInterop.release(sheet)
            ComInterop.release(worksheets)
            ComInterop.release(workbook)
            ComInterop.release(workbooks)
            ComInterop.release(excel)


class AccAuthSession(object):
    def __init__(self):
        self.AccessToken = ""
        self.RefreshToken = ""
        self.ExpiresAtUtc = DateTime.UtcNow
        self.ClientId = ""
        self.ClientSecret = ""

    @staticmethod
    def from_json(json_text):
        data = json.loads(json_text) if json_text else {}
        session = AccAuthSession()
        session.AccessToken = data.get("access_token", "")
        session.RefreshToken = data.get("refresh_token", "")
        expires_in = int(data.get("expires_in", 0))
        session.ExpiresAtUtc = DateTime.UtcNow.AddSeconds(expires_in if expires_in > 0 else 3000)
        return session


class AccAuthClient(object):
    def __init__(self, authorize_url, token_url, redirect_uri, scopes, log):
        self._authorize_url = authorize_url
        self._token_url = token_url
        self._redirect_uri = redirect_uri
        self._scopes = scopes
        self._log = log

    def authenticate(self, client_id, client_secret, timeout_seconds=180):
        state = Guid.NewGuid().ToString("N")
        auth_url = self._build_auth_url(client_id, state)

        self._log("Opening browser for login...")

        listener = HttpListener()
        listener.Prefixes.Add(self._redirect_uri)
        listener.Start()

        try:
            psi = ProcessStartInfo(auth_url)
            psi.UseShellExecute = True
            Process.Start(psi)

            async_result = listener.BeginGetContext(None, None)
            if not async_result.AsyncWaitHandle.WaitOne(TimeSpan.FromSeconds(timeout_seconds)):
                raise Exception("Login timed out. Please try again.")
            context = listener.EndGetContext(async_result)
            code = context.Request.QueryString["code"]
            returned_state = context.Request.QueryString["state"]
            error = context.Request.QueryString["error"]

            self._write_response(context.Response, error)

            if error:
                raise Exception("Auth error: {0}".format(error))
            if not code:
                raise Exception("Authorization code missing.")
            if state != returned_state:
                raise Exception("State mismatch. Possible tampering.")

            return self._exchange_code(client_id, client_secret, code)
        finally:
            listener.Stop()

    def refresh(self, client_id, client_secret, session):
        if not session or not session.RefreshToken:
            raise Exception("No refresh token available.")

        form = {
            "grant_type": "refresh_token",
            "refresh_token": session.RefreshToken,
        }

        headers = {
            "Authorization": "Basic " + self._basic_auth(client_id, client_secret)
        }
        body = build_form_body(form)
        json_text = http_send("POST", self._token_url, body, "application/x-www-form-urlencoded", headers)

        updated = AccAuthSession.from_json(json_text)
        session.AccessToken = updated.AccessToken
        session.RefreshToken = updated.RefreshToken
        session.ExpiresAtUtc = updated.ExpiresAtUtc

    def _build_auth_url(self, client_id, state):
        query = (
            "response_type=code"
            + "&client_id=" + Uri.EscapeDataString(client_id)
            + "&redirect_uri=" + Uri.EscapeDataString(self._redirect_uri)
            + "&scope=" + Uri.EscapeDataString(self._scopes)
            + "&state=" + Uri.EscapeDataString(state)
            + "&prompt=login"
        )
        return self._authorize_url + "?" + query

    def _exchange_code(self, client_id, client_secret, code):
        form = {
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": self._redirect_uri,
        }

        headers = {
            "Authorization": "Basic " + self._basic_auth(client_id, client_secret)
        }
        body = build_form_body(form)
        json_text = http_send("POST", self._token_url, body, "application/x-www-form-urlencoded", headers)
        return AccAuthSession.from_json(json_text)

    def _basic_auth(self, client_id, client_secret):
        raw = (client_id + ":" + client_secret)
        raw_bytes = Encoding.ASCII.GetBytes(raw)
        return Convert.ToBase64String(raw_bytes)

    def _write_response(self, response, error):
        message = (
            "You can close this window and return to Revit."
            if not error
            else "Authentication failed. You can close this window."
        )
        html = "<html><body><h3>{0}</h3></body></html>".format(message)
        buffer = Encoding.UTF8.GetBytes(html)
        response.ContentLength64 = buffer.Length
        stream = response.OutputStream
        stream.Write(buffer, 0, buffer.Length)
        stream.Close()


class AccDataClient(object):
    def __init__(self, auth_client, session, log):
        self._auth_client = auth_client
        self._session = session
        self._log = log

    def get_hubs(self):
        json_text = self._get("https://developer.api.autodesk.com/project/v1/hubs")
        root = json.loads(json_text) if json_text else {}
        result = []
        for item in root.get("data", []):
            attrs = item.get("attributes", {})
            result.append(HubInfo(item.get("id", ""), attrs.get("name", "")))
        return result

    def get_projects(self, hub_id):
        url = "https://developer.api.autodesk.com/project/v1/hubs/{0}/projects?page[limit]=200".format(
            Uri.EscapeDataString(hub_id)
        )
        result = []
        next_url = url
        while next_url:
            json_text = self._get(next_url)
            root = json.loads(json_text) if json_text else {}
            for item in root.get("data", []):
                attrs = item.get("attributes", {})
                result.append(ProjectInfo(item.get("id", ""), attrs.get("name", "")))
            next_url = get_next_link(root)
        return result

    def get_top_folders(self, hub_id, project_id):
        url = "https://developer.api.autodesk.com/project/v1/hubs/{0}/projects/{1}/topFolders".format(
            Uri.EscapeDataString(hub_id),
            Uri.EscapeDataString(project_id),
        )
        json_text = self._get(url)
        root = json.loads(json_text) if json_text else {}
        result = []
        for item in root.get("data", []):
            attrs = item.get("attributes", {})
            result.append(FolderNode(item.get("id", ""), attrs.get("displayName", "")))
        return result

    def get_folder_children(self, project_id, folder_id):
        children = []
        files = []
        url = "https://developer.api.autodesk.com/data/v1/projects/{0}/folders/{1}/contents".format(
            Uri.EscapeDataString(project_id),
            Uri.EscapeDataString(folder_id),
        )
        json_text = self._get(url)
        root = json.loads(json_text) if json_text else {}
        self._parse_folder_contents(root, children, files)
        return children

    def get_files_in_folder(self, project_id, folder_id):
        folders = []
        files = []
        url = "https://developer.api.autodesk.com/data/v1/projects/{0}/folders/{1}/contents".format(
            Uri.EscapeDataString(project_id),
            Uri.EscapeDataString(folder_id),
        )

        next_url = url
        while next_url:
            json_text = self._get(next_url)
            root = json.loads(json_text) if json_text else {}
            self._parse_folder_contents(root, folders, files)
            next_url = get_next_link(root)

        return files

    def update_file_description(self, project_id, item_id, description):
        self._ensure_token()
        url = "https://developer.api.autodesk.com/data/v1/projects/{0}/items/{1}".format(
            Uri.EscapeDataString(project_id),
            Uri.EscapeDataString(item_id),
        )

        body = json.dumps({
            "jsonapi": {"version": "1.0"},
            "data": {
                "type": "items",
                "id": item_id,
                "attributes": {
                    "extension": {
                        "data": {
                            "description": description,
                        }
                    }
                },
            },
        })

        headers = {"Authorization": "Bearer " + self._session.AccessToken}
        try:
            desc_preview = description if description is not None else ""
            if len(desc_preview) > 120:
                desc_preview = desc_preview[:117] + "..."
            self._log("PATCH item description: item_id={0} desc_len={1} desc='{2}'".format(
                item_id,
                len(description or ""),
                desc_preview,
            ))
            self._log("PATCH url: {0}".format(url))
            self._log("PATCH body: {0}".format(body))
        except Exception:
            pass
        http_send("PATCH", url, body, "application/vnd.api+json", headers)

    def update_version_description(self, project_id, version_id, description):
        self._ensure_token()
        url = "https://developer.api.autodesk.com/data/v1/projects/{0}/versions/{1}".format(
            Uri.EscapeDataString(project_id),
            Uri.EscapeDataString(version_id),
        )

        body = json.dumps({
            "jsonapi": {"version": "1.0"},
            "data": {
                "type": "versions",
                "id": version_id,
                "attributes": {
                    "description": description,
                },
            },
        })

        headers = {"Authorization": "Bearer " + self._session.AccessToken}
        try:
            desc_preview = description if description is not None else ""
            if len(desc_preview) > 120:
                desc_preview = desc_preview[:117] + "..."
            self._log("PATCH version description: version_id={0} desc_len={1} desc='{2}'".format(
                version_id,
                len(description or ""),
                desc_preview,
            ))
            self._log("PATCH url: {0}".format(url))
            self._log("PATCH body: {0}".format(body))
        except Exception:
            pass
        http_send("PATCH", url, body, "application/vnd.api+json", headers)

    def _get(self, url):
        self._ensure_token()
        headers = {"Authorization": "Bearer " + self._session.AccessToken}
        return http_send("GET", url, None, "application/json", headers)

    def get_item_detail(self, project_id, item_id):
        self._ensure_token()
        url = "https://developer.api.autodesk.com/data/v1/projects/{0}/items/{1}".format(
            Uri.EscapeDataString(project_id),
            Uri.EscapeDataString(item_id),
        )
        return self._get(url)

    def _ensure_token(self):
        if self._session.ExpiresAtUtc <= DateTime.UtcNow.AddMinutes(2):
            self._log("Refreshing token...")
            self._auth_client.refresh(self._session.ClientId, self._session.ClientSecret, self._session)

    def _parse_folder_contents(self, root, folders, files):
        for item in root.get("data", []):
            item_type = item.get("type", "")
            attrs = item.get("attributes", {})

            if item_type.lower() == "folders":
                folders.append(FolderNode(item.get("id", ""), attrs.get("displayName", "")))
            elif item_type.lower() == "items":
                ext = attrs.get("extension", {})
                ext_type = ext.get("type", "")
                can_update = True
                files.append(FileItem(
                    item.get("id", ""),
                    attrs.get("displayName", ""),
                    attrs.get("description", ""),
                    ext_type,
                    can_update,
                ))


class HubInfo(object):
    def __init__(self, hub_id, name):
        self.Id = hub_id
        self.Name = name
    def __str__(self):
        return self.Name


class ProjectInfo(object):
    def __init__(self, project_id, name):
        self.Id = project_id
        self.Name = name
    def __str__(self):
        return self.Name


class FolderNode(object):
    def __init__(self, folder_id, name, is_placeholder=False, is_up_level=False):
        self.Id = folder_id
        self.Name = name
        self.IsPlaceholder = is_placeholder
        self.IsUpLevel = is_up_level
        self.Children = ObservableCollection[Object]()
        self.IsLoaded = False
        self.IsLoading = False
        if not is_placeholder and not is_up_level:
            self.Children.Add(FolderNode("", "", True))


class FileItem(object):
    def __init__(self, item_id, display_name, description, ext_type, can_update):
        self.Id = item_id
        self.DisplayName = display_name
        self.Description = description
        self.ExtensionType = ext_type
        self.CanUpdateDescription = can_update


def load_window():
    xaml_path = os.path.join(os.path.dirname(__file__), "ui.xaml")
    if not os.path.exists(xaml_path):
        UI.TaskDialog.Show(TITLE, "ui.xaml not found. Ensure ui.xaml is next to script.py.")
        return None

    stream = None
    try:
        stream = FileStream(xaml_path, FileMode.Open, FileAccess.Read)
        return XamlReader.Load(stream)
    finally:
        if stream is not None:
            stream.Close()


class AccDocsWindow(object):
    def __init__(self):
        self._window = load_window()
        if self._window is None:
            return

        self.client_id_box = self._window.FindName("ClientIdBox")
        self.client_secret_box = self._window.FindName("ClientSecretBox")
        self.redirect_uri_box = self._window.FindName("RedirectUriBox")
        self.login_button = self._window.FindName("LoginButton")
        self.hub_combo = self._window.FindName("HubCombo")
        self.project_combo = self._window.FindName("ProjectCombo")
        self.folder_url_box = self._window.FindName("FolderUrlBox")
        self.load_folder_button = self._window.FindName("LoadFolderButton")
        self.folder_tree = self._window.FindName("FolderTree")
        self.refresh_folders_button = self._window.FindName("RefreshFoldersButton")
        self.refresh_files_button = self._window.FindName("RefreshFilesButton")
        self.file_list = self._window.FindName("FileList")
        self.excel_path_box = self._window.FindName("ExcelPathBox")
        self.excel_status_text = self._window.FindName("ExcelStatusText")
        self.browse_excel_button = self._window.FindName("BrowseExcelButton")
        self.apply_button = self._window.FindName("ApplyButton")
        self.log_box = self._window.FindName("LogBox")
        self.status_text = self._window.FindName("StatusText")

        self.redirect_uri_box.Text = REDIRECT_URI

        self.hub_combo.IsEnabled = False
        self.project_combo.IsEnabled = False
        self.folder_tree.IsEnabled = False
        self.refresh_folders_button.IsEnabled = False
        self.browse_excel_button.IsEnabled = False
        self.apply_button.IsEnabled = False

        self._folder_nodes = ObservableCollection[Object]()
        self._file_items = ObservableCollection[Object]()
        self._folder_nav_stack = []
        self.folder_tree.ItemsSource = self._folder_nodes
        self.file_list.ItemsSource = self._file_items

        self._session = None
        self._auth_client = None
        self._data_client = None
        self._excel_rows = None
        self._auth_in_progress = False
        self._last_selected_folder_id = None
        self._config = pyrevit_script.get_config()

        self.login_button.Click += self.on_sign_in
        self.hub_combo.SelectionChanged += self.on_hub_changed
        self.project_combo.SelectionChanged += self.on_project_changed
        self.folder_tree.SelectedItemChanged += self.on_folder_selected
        self.folder_tree.AddHandler(TreeViewItem.ExpandedEvent, RoutedEventHandler(self.on_folder_expanded))
        self.folder_tree.MouseDoubleClick += MouseButtonEventHandler(self.on_folder_double_click)
        self.refresh_folders_button.Click += self.on_refresh_folders
        if self.refresh_files_button:
            self.refresh_files_button.Click += self.on_refresh_files
        self.browse_excel_button.Click += self.on_browse_excel
        self.apply_button.Click += self.on_apply_descriptions
        if self.load_folder_button:
            self.load_folder_button.Click += self.on_load_folder_url

        self.log("Ready.")

        try:
            creds = _get_env_credentials()
            if creds.get(ENV_CLIENT_ID_KEY):
                self.client_id_box.Text = creds.get(ENV_CLIENT_ID_KEY)
            if creds.get(ENV_CLIENT_SECRET_KEY):
                self.client_secret_box.Password = creds.get(ENV_CLIENT_SECRET_KEY)
        except Exception:
            pass

        try:
            self._restore_cached_inputs()
            self._restore_cached_session()
        except Exception:
            pass

        try:
            self._window.Closing += self.on_window_closing
        except Exception:
            pass

    def _as_collection(self, items):
        col = ObservableCollection[Object]()
        if items:
            for item in items:
                col.Add(item)
        return col

    def show(self):
        if self._window is None:
            return
        self._window.ShowDialog()

    def on_window_closing(self, sender, args):
        try:
            if self.folder_url_box and self.folder_url_box.Text:
                self._config.last_folder_url = self.folder_url_box.Text
            if self.excel_path_box and self.excel_path_box.Text:
                self._config.last_excel_path = self.excel_path_box.Text
            pyrevit_script.save_config()
        except Exception:
            pass

    def _restore_cached_inputs(self):
        try:
            last_folder_url = getattr(self._config, "last_folder_url", None)
            if last_folder_url and self.folder_url_box:
                self.folder_url_box.Text = last_folder_url
        except Exception:
            pass

        try:
            last_excel_path = getattr(self._config, "last_excel_path", None)
            if last_excel_path and self.excel_path_box:
                self.excel_path_box.Text = last_excel_path
                if os.path.exists(last_excel_path):
                    try:
                        self._excel_rows = ExcelReader.read_excel(last_excel_path)
                        unique_rows = set()
                        try:
                            unique_rows = set([x.FileName for x in self._excel_rows.values() if x])
                        except Exception:
                            unique_rows = set()
                        self.excel_status_text.Text = "Loaded {0} rows from Excel (case-insensitive; extension optional).".format(len(unique_rows))
                        self.log("Excel loaded: {0}".format(last_excel_path))
                        self._log_excel_diagnostics()
                    except Exception as ex:
                        self._excel_rows = None
                        self.excel_status_text.Text = "Excel load failed."
                        self.log("Excel load failed: {0}".format(ex))
        except Exception:
            pass

    def _restore_cached_session(self):
        try:
            token = getattr(self._config, "acc_token", None)
            expires_ticks = getattr(self._config, "acc_token_expires", None)
            if not token or not expires_ticks:
                return

            try:
                expires_ticks = int(expires_ticks)
            except Exception:
                return

            expires_at = DateTime(expires_ticks)
            if expires_at <= DateTime.UtcNow.AddMinutes(2):
                try:
                    self._config.acc_token = None
                    self._config.acc_token_expires = None
                    pyrevit_script.save_config()
                except Exception:
                    pass
                return

            session = AccAuthSession()
            session.AccessToken = token
            session.ExpiresAtUtc = expires_at
            session.ClientId = (self.client_id_box.Text or "").strip()
            session.ClientSecret = (self.client_secret_box.Password or "").strip()

            auth_client = AccAuthClient(OAUTH_AUTHORIZE_URL, OAUTH_TOKEN_URL, REDIRECT_URI, DEFAULT_SCOPES, self.log)
            data_client = AccDataClient(auth_client, session, self.log)

            self._auth_client = auth_client
            self._session = session
            self._data_client = data_client

            self.status_text.Text = "Signed in (cached)"
            self.log("Using cached token.")
            self.hub_combo.IsEnabled = True
            self.project_combo.IsEnabled = True
            self.folder_tree.IsEnabled = True
            self.refresh_folders_button.IsEnabled = True
            self.browse_excel_button.IsEnabled = True
            self.apply_button.IsEnabled = True

            self._load_hubs_async()
        except Exception:
            pass

    def _load_hubs_async(self):
        def worker():
            try:
                hubs = self._data_client.get_hubs()

                def ui_success():
                    self.hub_combo.ItemsSource = self._as_collection(hubs)
                    self.hub_combo.DisplayMemberPath = "Name"
                    self.hub_combo.SelectedValuePath = "Id"
                    if hubs:
                        self.hub_combo.SelectedIndex = 0
                    self.status_text.Text = "Hubs loaded"
                self._window.Dispatcher.Invoke(Action(ui_success))
            except Exception as ex:
                def ui_fail():
                    self.log("Failed to load hubs: {0}".format(ex))
                    self.status_text.Text = "Hub load failed"
                self._window.Dispatcher.Invoke(Action(ui_fail))

        t = Thread(ThreadStart(worker))
        t.IsBackground = True
        t.Start()

    def _log_excel_diagnostics(self):
        diag = ExcelReader.LastDiagnostics
        if not diag:
            return
        try:
            self.log("Excel diagnostics: range R{0}-R{1}, C{2}-C{3}, desc_col={4}, rows={5}".format(
                diag.get("row_start"),
                diag.get("row_end"),
                diag.get("col_start"),
                diag.get("col_end"),
                diag.get("description_index"),
                diag.get("rows_read"),
            ))
        except Exception:
            pass
        try:
            sheet_name = diag.get("sheet")
            if sheet_name:
                self.log("Excel sheet: {0} (rows={1}, desc_rows={2})".format(
                    sheet_name,
                    diag.get("sheet_rows"),
                    diag.get("sheet_desc_rows"),
                ))
        except Exception:
            pass
        try:
            headers = diag.get("headers") or []
            if headers:
                self.log("Excel headers (row 1): {0}".format(", ".join(headers)))
        except Exception:
            pass
        try:
            samples = diag.get("samples") or []
            for idx, sample in enumerate(samples):
                file_val = sample[0] if len(sample) > 0 else ""
                desc_val = sample[1] if len(sample) > 1 else ""
                self.log("Excel sample {0}: file='{1}' desc='{2}'".format(idx + 1, file_val, desc_val))
        except Exception:
            pass
        try:
            rows = diag.get("rows") or []
            if rows:
                self.log("Excel rows (A,B): showing up to {0}".format(len(rows)))
            for idx, row in enumerate(rows):
                file_val = row[0] if len(row) > 0 else ""
                desc_val = row[1] if len(row) > 1 else ""
                self.log("Excel row {0}: file='{1}' desc='{2}'".format(idx + 1, file_val, desc_val))
        except Exception:
            pass

    def on_sign_in(self, sender, args):
        client_id = (self.client_id_box.Text or "").strip()
        client_secret = (self.client_secret_box.Password or "").strip()

        if not client_id or not client_secret:
            MessageBox.Show(
                "Client ID and Client Secret are required. Set them in .env as CLIENT_ID and CLIENT_SECRET.",
                TITLE,
                MessageBoxButton.OK,
                MessageBoxImage.Warning,
            )
            return


        if self._auth_in_progress:
            return
        self._auth_in_progress = True
        self.login_button.IsEnabled = False
        self.status_text.Text = "Signing in..."

        def worker():
            try:
                auth_client = AccAuthClient(OAUTH_AUTHORIZE_URL, OAUTH_TOKEN_URL, REDIRECT_URI, DEFAULT_SCOPES, self.log)
                session = auth_client.authenticate(client_id, client_secret)
                session.ClientId = client_id
                session.ClientSecret = client_secret
                data_client = AccDataClient(auth_client, session, self.log)
                hubs = data_client.get_hubs()

                def ui_success():
                    self._auth_client = auth_client
                    self._session = session
                    self._data_client = data_client

                    self.status_text.Text = "Signed in"
                    self.log("Signed in successfully.")
                    self.hub_combo.IsEnabled = True
                    self.project_combo.IsEnabled = True
                    self.folder_tree.IsEnabled = True
                    self.refresh_folders_button.IsEnabled = True
                    self.browse_excel_button.IsEnabled = True
                    self.apply_button.IsEnabled = True

                    self.hub_combo.ItemsSource = self._as_collection(hubs)
                    self.hub_combo.DisplayMemberPath = "Name"
                    self.hub_combo.SelectedValuePath = "Id"
                    if hubs:
                        self.hub_combo.SelectedIndex = 0
                    else:
                        self.status_text.Text = "No hubs found (check app provisioning)"

                    self.login_button.IsEnabled = True
                    self._auth_in_progress = False
                    try:
                        self._config.acc_token = session.AccessToken
                        self._config.acc_token_expires = session.ExpiresAtUtc.Ticks
                        pyrevit_script.save_config()
                    except Exception:
                        pass

                self._window.Dispatcher.Invoke(Action(ui_success))
            except Exception as ex:
                def ui_fail():
                    self.status_text.Text = "Sign in failed"
                    self.log("Sign in failed: {0}".format(ex))
                    self.login_button.IsEnabled = True
                    self._auth_in_progress = False
                self._window.Dispatcher.Invoke(Action(ui_fail))

        t = Thread(ThreadStart(worker))
        t.IsBackground = True
        t.Start()

    def load_hubs(self):
        try:
            self.status_text.Text = "Loading hubs..."
            hubs = self._data_client.get_hubs()
            self.hub_combo.ItemsSource = self._as_collection(hubs)
            self.hub_combo.DisplayMemberPath = "Name"
            self.hub_combo.SelectedValuePath = "Id"
            if hubs:
                self.hub_combo.SelectedIndex = 0
            self.status_text.Text = "Hubs loaded"
        except Exception as ex:
            self.log("Failed to load hubs: {0}".format(ex))
            self.status_text.Text = "Hub load failed"

    def on_hub_changed(self, sender, args):
        hub = self.hub_combo.SelectedItem
        if hub is None:
            return

        try:
            self.status_text.Text = "Loading projects..."
            projects = self._data_client.get_projects(hub.Id)
            try:
                projects = sorted(projects, key=lambda p: (p.Name or "").lower())
            except Exception:
                pass
            self.project_combo.ItemsSource = self._as_collection(projects)
            self.project_combo.DisplayMemberPath = "Name"
            self.project_combo.SelectedValuePath = "Id"
            if projects:
                self.project_combo.SelectedIndex = 0
            self.status_text.Text = "Projects loaded"
        except Exception as ex:
            self.log("Failed to load projects: {0}".format(ex))
            self.status_text.Text = "Project load failed"

    def on_project_changed(self, sender, args):
        self.load_top_folders()

    def load_top_folders(self):
        hub = self.hub_combo.SelectedItem
        project = self.project_combo.SelectedItem
        if hub is None or project is None:
            return

        try:
            self.status_text.Text = "Loading folders..."
            self._folder_nodes.Clear()
            self._file_items.Clear()
            self._folder_nav_stack = []
            self._last_selected_folder_id = None

            top_folders = self._data_client.get_top_folders(hub.Id, project.Id)
            for folder in top_folders:
                try:
                    if folder.Children is not None and folder.Children.Count == 0:
                        folder.Children.Add(FolderNode("", "", True))
                except Exception:
                    pass
                self._folder_nodes.Add(folder)
                try:
                    count = folder.Children.Count if folder.Children is not None else 0
                    self.log("Top folder '{0}' children={1}".format(folder.Name, count))
                except Exception:
                    pass

            # Auto-expand Project Files if available
            for folder in top_folders:
                try:
                    if folder.Name and folder.Name.lower() == "project files":
                        self._expand_folder_node(project.Id, folder)
                        break
                except Exception:
                    continue

            self.status_text.Text = "Folders loaded"
        except Exception as ex:
            self.log("Failed to load folders: {0}".format(ex))
            self.status_text.Text = "Folder load failed"

    def on_folder_expanded(self, sender, args):
        item = args.OriginalSource
        node = item.DataContext if item else None
        if node is None or node.IsLoaded or node.IsPlaceholder:
            return

        project = self.project_combo.SelectedItem
        if project is None:
            return

        try:
            self.log("Loading subfolders for '{0}'...".format(node.Name))
            node.IsLoading = True
            self._expand_folder_node(project.Id, node)
            try:
                count = node.Children.Count if node.Children is not None else 0
                self.log("Loaded {0} subfolder(s) under '{1}'.".format(count, node.Name))
            except Exception:
                pass
        except Exception as ex:
            self.log("Failed to expand folder: {0}".format(ex))
        finally:
            node.IsLoading = False

    def _expand_folder_node(self, project_id, node):
        children = self._data_client.get_folder_children(project_id, node.Id)
        node.Children.Clear()
        for child in children:
            node.Children.Add(child)
        node.IsLoaded = True

    def on_folder_selected(self, sender, args):
        node = args.NewValue
        if node is None:
            return
        if getattr(node, "IsUpLevel", False):
            return

        project = self.project_combo.SelectedItem
        if project is None:
            return

        try:
            try:
                count = node.Children.Count if node.Children is not None else 0
                self.log("Folder selected '{0}' children={1}".format(node.Name, count))
            except Exception:
                pass
            if self._last_selected_folder_id == node.Id:
                try:
                    self.log("Reselected same folder; refreshing files.")
                except Exception:
                    pass
            self._last_selected_folder_id = node.Id
            self._load_files_for_folder(project.Id, node)
        except Exception as ex:
            self.log("Failed to load files: {0}".format(ex))
            self.status_text.Text = "File load failed"

    def on_folder_double_click(self, sender, args):
        try:
            node = self.folder_tree.SelectedItem
        except Exception:
            node = None
        if node is None:
            return
        if getattr(node, "IsUpLevel", False):
            try:
                if self._folder_nav_stack:
                    previous = self._folder_nav_stack.pop()
                    self._folder_nodes.Clear()
                    for item in previous:
                        self._folder_nodes.Add(item)
                    self.status_text.Text = "Folders loaded"
                    self.log("Returned to previous folder level.")
            except Exception as ex:
                self.log("Failed to return to previous level: {0}".format(ex))
            return

        project = self.project_combo.SelectedItem
        if project is None:
            return

        try:
            self.log("Double-clicked folder '{0}' - loading subfolders...".format(node.Name))
        except Exception:
            pass

        try:
            if not node.IsLoaded:
                self._expand_folder_node(project.Id, node)
            try:
                children = [c for c in node.Children if not getattr(c, "IsPlaceholder", False)]
            except Exception:
                children = []
            try:
                self.log("Loaded {0} subfolder(s) under '{1}'.".format(len(children), node.Name))
            except Exception:
                pass
            if children:
                try:
                    self._folder_nav_stack.append(list(self._folder_nodes))
                    self._folder_nodes.Clear()
                    up = FolderNode("", "..", False, True)
                    try:
                        up.Children.Clear()
                    except Exception:
                        pass
                    self._folder_nodes.Add(up)
                    for child in children:
                        self._folder_nodes.Add(child)
                    self.status_text.Text = "Folders loaded"
                except Exception as ex:
                    self.log("Failed to show subfolders for '{0}': {1}".format(getattr(node, "Name", ""), ex))
            else:
                self.log("No subfolders under '{0}'.".format(getattr(node, "Name", "")))
        except Exception as ex:
            self.log("Failed to load subfolders for '{0}': {1}".format(getattr(node, "Name", ""), ex))

    def on_refresh_folders(self, sender, args):
        self.load_top_folders()

    def on_refresh_files(self, sender, args):
        project = self.project_combo.SelectedItem
        if project is None:
            return
        try:
            node = self.folder_tree.SelectedItem
        except Exception:
            node = None
        if node is None or getattr(node, "IsUpLevel", False):
            return
        try:
            self.log("Refreshing files for '{0}'...".format(node.Name))
        except Exception:
            pass
        self._load_files_for_folder(project.Id, node)

    def _load_files_for_folder(self, project_id, node):
        try:
            self.status_text.Text = "Loading files..."
            self._file_items.Clear()
            files = self._data_client.get_files_in_folder(project_id, node.Id)
            for item in files:
                self._file_items.Add(item)
            if files:
                self.status_text.Text = "Files loaded"
            else:
                self.status_text.Text = "No files in this folder. Select a subfolder."
        except Exception as ex:
            self.log("Failed to load files: {0}".format(ex))
            self.status_text.Text = "File load failed"

    def on_load_folder_url(self, sender, args):
        if not self._data_client:
            MessageBox.Show("Sign in first.", TITLE, MessageBoxButton.OK, MessageBoxImage.Warning)
            return
        url = (self.folder_url_box.Text or "").strip() if self.folder_url_box else ""
        project_id, folder_id = self._parse_folder_url(url)
        if not folder_id:
            MessageBox.Show("Invalid folder URL. Paste a Project Files folder URL from ACC.", TITLE, MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        candidates = []
        if project_id:
            candidates.extend(self._normalize_project_id(project_id))

        project = self.project_combo.SelectedItem
        if project and project.Id not in candidates:
            candidates.append(project.Id)

        if not candidates:
            MessageBox.Show("Project ID not found in URL. Select a project from the list and try again.", TITLE, MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        try:
            self._config.last_folder_url = url
            pyrevit_script.save_config()
        except Exception:
            pass

        if self.load_folder_button:
            self.load_folder_button.IsEnabled = False
        self.status_text.Text = "Loading files from URL..."

        def worker():
            last_error = None
            for pid in candidates:
                try:
                    files = self._data_client.get_files_in_folder(pid, folder_id)
                    def ui_ok():
                        self._file_items.Clear()
                        for item in files:
                            self._file_items.Add(item)
                        if files:
                            self.status_text.Text = "Files loaded from URL"
                        else:
                            self.status_text.Text = "No files in this folder."
                        if self.load_folder_button:
                            self.load_folder_button.IsEnabled = True
                    self._window.Dispatcher.Invoke(Action(ui_ok))
                    return
                except Exception as ex:
                    last_error = ex
                    if "Invalid project id" in str(ex):
                        continue
                    break

            def ui_fail():
                self.status_text.Text = "Folder load failed"
                self.log("Failed to load folder URL: {0}".format(last_error))
                if self.load_folder_button:
                    self.load_folder_button.IsEnabled = True
            self._window.Dispatcher.Invoke(Action(ui_fail))

        t = Thread(ThreadStart(worker))
        t.IsBackground = True
        t.Start()

    def on_browse_excel(self, sender, args):
        dialog = OpenFileDialog()
        dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|Excel Files (*.xls)|*.xls"
        dialog.Multiselect = False
        if not dialog.ShowDialog():
            return

        self.excel_path_box.Text = dialog.FileName
        try:
            self._excel_rows = ExcelReader.read_excel(dialog.FileName)
            unique_rows = set()
            try:
                unique_rows = set([x.FileName for x in self._excel_rows.values() if x])
            except Exception:
                unique_rows = set()
            self.excel_status_text.Text = "Loaded {0} rows from Excel (case-insensitive; extension optional).".format(len(unique_rows))
            self.log("Excel loaded: {0}".format(dialog.FileName))
            self._log_excel_diagnostics()
            try:
                self._config.last_excel_path = dialog.FileName
                pyrevit_script.save_config()
            except Exception:
                pass
        except Exception as ex:
            self._excel_rows = None
            self.excel_status_text.Text = "Excel load failed."
            self.log("Excel load failed: {0}".format(ex))

    def on_apply_descriptions(self, sender, args):
        if not self._excel_rows:
            MessageBox.Show("Load an Excel file first.", TITLE, MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        project = self.project_combo.SelectedItem
        if project is None:
            return

        updated = 0
        skipped = 0
        try:
            keys_sample = []
            for key in self._excel_rows.keys():
                keys_sample.append(key)
                if len(keys_sample) >= 5:
                    break
            if keys_sample:
                self.log("Excel keys sample: {0}".format(", ".join(keys_sample)))
        except Exception:
            pass

        for file_item in list(self._file_items):
            item_label = "{0} [{1}]".format(file_item.DisplayName, file_item.ExtensionType)
            if not file_item.Id:
                skipped += 1
                self.log("Skipped (missing item id): {0}".format(item_label))
                continue

            row_key = ExcelReader._normalize_name(file_item.DisplayName)
            row = self._excel_rows.get(row_key)
            if not row:
                row_key_base = ExcelReader._normalize_base(file_item.DisplayName)
                row = self._excel_rows.get(row_key_base)
            if not row:
                skipped += 1
                self.log("Skipped (no Excel row): {0} | key='{1}' base='{2}'".format(item_label, row_key, row_key_base))
                continue

            if not row.Description:
                skipped += 1
                self.log("Skipped (empty description): {0} | excel='{1}'".format(item_label, row.FileName))
                continue

            try:
                try:
                    desc_preview = row.Description if row.Description is not None else ""
                    if len(desc_preview) > 120:
                        desc_preview = desc_preview[:117] + "..."
                    self.log("Updating: {0} | type={1} | item_id={2} | desc_len={3} | desc='{4}'".format(
                        file_item.DisplayName,
                        file_item.ExtensionType,
                        file_item.Id,
                        len(row.Description or ""),
                        desc_preview,
                    ))
                except Exception:
                    pass
                try:
                    item_json = self._data_client.get_item_detail(project.Id, file_item.Id)
                    item_root = json.loads(item_json) if item_json else {}
                    item_data = item_root.get("data", {}) if isinstance(item_root, dict) else {}
                    item_attrs = item_data.get("attributes", {}) if isinstance(item_data, dict) else {}
                    ext = item_attrs.get("extension", {}) if isinstance(item_attrs, dict) else {}
                    ext_type = ext.get("type", "")
                    ext_data = ext.get("data", {}) if isinstance(ext, dict) else {}
                    current_desc = ext_data.get("description", "")
                    tip = None
                    try:
                        rel = item_data.get("relationships", {}) if isinstance(item_data, dict) else {}
                        tip_rel = rel.get("tip", {}) if isinstance(rel, dict) else {}
                        tip_data = tip_rel.get("data", {}) if isinstance(tip_rel, dict) else {}
                        tip = tip_data.get("id", None)
                    except Exception:
                        tip = None
                    self.log("Pre-update item: ext_type='{0}' desc='{1}' tip='{2}'".format(ext_type, current_desc, tip or ""))
                except Exception as ex:
                    self.log("Pre-update fetch failed: {0}".format(ex))
                    item_data = {}
                    tip = None

                if file_item.ExtensionType == "items:autodesk.bim360:C4RModel":
                    skipped += 1
                    self.log("Skipped (C4RModel not supported for description update): {0}".format(item_label))
                    continue

                if not tip:
                    skipped += 1
                    self.log("Skipped (missing tip version id): {0}".format(item_label))
                    continue

                self._data_client.update_version_description(project.Id, tip, row.Description)
                file_item.Description = row.Description
                updated += 1
                self.log("Updated: {0}".format(item_label))
            except Exception as ex:
                self.log("Failed to update {0}: {1}".format(item_label, ex))

        try:
            self.file_list.Items.Refresh()
        except Exception:
            pass

        self.status_text.Text = "Update complete"
        self.log("Update finished. Updated: {0}, Skipped: {1}".format(updated, skipped))

    def _parse_folder_url(self, url):
        if not url:
            return None, None
        project_id = None
        folder_id = None

        try:
            uri = Uri(url)
            if uri and uri.Query:
                query = uri.Query.lstrip("?")
                parts = query.split("&")
                for part in parts:
                    if part.startswith("folderUrn="):
                        folder_id = Uri.UnescapeDataString(part.split("=", 1)[1])
                    if part.startswith("projectId="):
                        project_id = Uri.UnescapeDataString(part.split("=", 1)[1])
        except Exception:
            pass

        m = re.search(r"/projects/([a-zA-Z0-9\-\.]+)", url)
        if m:
            project_id = m.group(1)

        if not folder_id:
            m = re.search(r"/folders/([^/?#]+)", url)
            if m:
                folder_id = Uri.UnescapeDataString(m.group(1))

        if folder_id and folder_id.startswith("urn%3A"):
            try:
                folder_id = Uri.UnescapeDataString(folder_id)
            except Exception:
                pass

        return project_id, folder_id

    def _normalize_project_id(self, project_id):
        ids = []
        if not project_id:
            return ids
        pid = project_id.strip()
        if pid.startswith("urn:"):
            try:
                tail = pid.split(":")[-1]
            except Exception:
                tail = pid
            pid = tail or pid
        ids.append(pid)
        if not (pid.startswith("b.") or pid.startswith("a.")):
            ids.append("b." + pid)
            ids.append("a." + pid)
        # de-dupe
        result = []
        for x in ids:
            if x and x not in result:
                result.append(x)
        return result

    def log(self, message):
        if self.log_box is None:
            return
        try:
            if self._window is not None and not self._window.Dispatcher.CheckAccess():
                def _ui_log():
                    self.log(message)
                self._window.Dispatcher.Invoke(Action(_ui_log))
                return
        except Exception:
            pass
        stamp = DateTime.Now.ToString("HH:mm:ss")
        self.log_box.AppendText("{0}  {1}\n".format(stamp, message))
        self.log_box.ScrollToEnd()


def build_form_body(form):
    parts = []
    for key, value in form.items():
        parts.append(Uri.EscapeDataString(key) + "=" + Uri.EscapeDataString(value or ""))
    return "&".join(parts)


def http_send(method, url, body, content_type, headers):
    request = WebRequest.Create(url)
    request.Method = method
    request.Accept = "application/json, application/vnd.api+json"
    request.UserAgent = "WWP BIM Tools"

    if content_type:
        request.ContentType = content_type

    if headers:
        for key, value in headers.items():
            request.Headers[key] = value

    if body:
        data = Encoding.UTF8.GetBytes(body)
        request.ContentLength = data.Length
        stream = request.GetRequestStream()
        stream.Write(data, 0, data.Length)
        stream.Close()
    elif method in ("POST", "PATCH"):
        request.ContentLength = 0

    try:
        response = request.GetResponse()
        reader = StreamReader(response.GetResponseStream())
        text = reader.ReadToEnd()
        reader.Close()
        response.Close()
        return text
    except WebException as ex:
        if ex.Response:
            reader = StreamReader(ex.Response.GetResponseStream())
            error_text = reader.ReadToEnd()
            reader.Close()
            raise Exception(error_text)
        raise


def get_next_link(root):
    links = root.get("links", {}) if isinstance(root, dict) else {}
    next_link = links.get("next") if isinstance(links, dict) else None
    if isinstance(next_link, dict):
        return next_link.get("href")
    return None


try:
    window = AccDocsWindow()
    window.show()
except Exception:
    UI.TaskDialog.Show("WWP BIM Tools", "Unexpected error:\n" + traceback.format_exc())
