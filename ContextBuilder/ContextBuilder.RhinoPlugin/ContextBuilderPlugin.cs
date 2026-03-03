using System;
using Rhino.PlugIns;

namespace ContextBuilder.RhinoPlugin;

public sealed class ContextBuilderPlugin : PlugIn
{
    public static ContextBuilderPlugin Instance { get; private set; } = null!;

    public ContextBuilderPlugin()
    {
        Instance = this;
    }
}
