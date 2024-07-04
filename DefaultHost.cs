using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Globalization;
using System.Management.Automation.Host;

namespace DefaultHost : PSHost
{
    private CultureInfo _currentCulture;
private CultureInfo _currentUICulture;
private InformationDelegate _onInformation;

/// <summary>
/// Initializes a new instance of the <see cref="DefaultHost"/> class.
/// </summary>
/// <param name="currentCulture">The current culture.</param>
/// <param name="currentUICulture">The current UI culture.</param>
/// <exception cref="ArgumentNullException">Thrown if <paramref name="currentCulture"/> is null.</exception>
/// <exception cref="ArgumentNullException">Thrown if <paramref name="currentUICulture"/> is null.</exception>
public DefaultHost(CultureInfo currentCulture, CultureInfo currentUICulture)
{
    _currentCulture = currentCulture ?? throw new ArgumentNullException(nameof(currentCulture));
    _currentUICulture = currentUICulture ?? throw new ArgumentNullException(nameof(currentUICulture));
}

public event InformationDelegate OnInformation
{
    [MethodImpl(MethodImplOptions.Synchronized)]
    add
    {
        _onInformation = (InformationDelegate)Delegate.Combine(_onInformation, value);
        GC.KeepAlive(this);
    }
    [MethodImpl(MethodImplOptions.Synchronized)]
    remove
    {
        _onInformation = (InformationDelegate)Delegate.Remove(_onInformation, value);
        GC.KeepAlive(this);
    }
}

protected void RaiseOnInformation(string value)
{
    _onInformation?.Invoke(value);
}

private void HandleInformation(string information)
{
    RaiseOnInformation(information);
}

public override string Name => "Default Host";

public override Version Version => new Version(1, 0);

public override Guid InstanceId => Guid.NewGuid();

public override PSHostUserInterface UI
{
    get
    {
        var defaultHostUserInterface = new DefaultHostUserInterface();
        defaultHostUserInterface.OnInformation += HandleInformation;
        GC.KeepAlive(this);
        return defaultHostUserInterface;
    }
}

public override CultureInfo CurrentCulture
{
    get => _currentCulture;
    set => _currentCulture = value ?? throw new ArgumentNullException(nameof(value));
}

public override CultureInfo CurrentUICulture
{
    get => _currentUICulture;
    set => _currentUICulture = value ?? throw new ArgumentNullException(nameof(value));
}

public override void SetShouldExit(int exitCode)
{
    throw new NotSupportedException("Not implemented.");
}

public override void EnterNestedPrompt()
{
    throw new NotSupportedException("Nested prompt attempted.");
}

public override void ExitNestedPrompt()
{
    throw new NotSupportedException("Not implemented.");
}

public override void NotifyBeginApplication()
{
    throw new NotSupportedException("Not implemented.");
}

public override void NotifyEndApplication()
{
    throw new NotSupportedException("Exit prompt attempted.");
}

public delegate void InformationDelegate(string information);
}
