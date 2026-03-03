using Rhino;
using Rhino.Commands;
using ContextBuilder.RhinoPlugin.UI;

namespace ContextBuilder.RhinoPlugin.Commands;

public sealed class ContextBuilderCommand : Command
{
    public override string EnglishName => "ContextBuilderDownload";

    protected override Result RunCommand(RhinoDoc doc, RunMode mode)
    {
        var dialog = new ContextBuilderDialog(doc);
        dialog.ShowModal();
        return Result.Success;
    }
}
