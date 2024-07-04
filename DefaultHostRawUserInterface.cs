using System;
using System.Management.Automation.Host;
using System.Runtime.InteropServices;

public class DefaultHostRawUserInterface : PSHostRawUserInterface
{
    public override ConsoleColor ForegroundColor
    {
        get => Console.ForegroundColor;
        set => Console.ForegroundColor = value;
    }

    public override ConsoleColor BackgroundColor
    {
        get => Console.BackgroundColor;
        set => Console.BackgroundColor = value;
    }

    public override Coordinates CursorPosition
    {
        get => new Coordinates(Console.CursorLeft, Console.CursorTop);
        set => Console.SetCursorPosition(value.X, value.Y);
    }

    public override Coordinates WindowPosition
    {
        get => new Coordinates(Console.WindowLeft, Console.WindowTop);
        set => Console.SetWindowPosition(value.X, value.Y);
    }

    public override int CursorSize
    {
        get => Console.CursorSize;
        set => Console.CursorSize = value;
    }

    public override Size BufferSize
    {
        get => new Size(Console.BufferWidth, Console.BufferHeight);
        set => Console.SetBufferSize(value.Width, value.Height);
    }

    public override Size WindowSize
    {
        get => new Size(Console.WindowWidth, Console.WindowHeight);
        set => Console.SetWindowSize(value.Width, value.Height);
    }

    public override Size MaxWindowSize
    {
        get => new Size(Console.LargestWindowWidth, Console.LargestWindowHeight);
    }

    public override Size MaxPhysicalWindowSize
    {
        get => new Size(Console.LargestWindowWidth, Console.LargestWindowHeight);
    }

    public override bool KeyAvailable => Console.KeyAvailable;

    public override string WindowTitle
    {
        get => Console.Title;
        set => Console.Title = value;
    }

    public override KeyInfo ReadKey(ReadKeyOptions options)
    {
        var intercept = (options & ReadKeyOptions.IncludeKeyDown) != 0;
        var keyInfo = Console.ReadKey(intercept);
        return new KeyInfo(keyInfo.KeyChar, (int)keyInfo.Key, (ControlKeyStates)0, keyInfo.Key == ConsoleKey.Escape);
    }

    public override void FlushInputBuffer()
    {
        while (Console.KeyAvailable)
        {
            Console.ReadKey(true);
        }
    }

    public override void SetBufferContents(Rectangle rectangle, BufferCell fill)
    {
        for (int y = rectangle.Top; y <= rectangle.Bottom; y++)
        {
            for (int x = rectangle.Left; x <= rectangle.Right; x++)
            {
                Console.SetCursorPosition(x, y);
                Console.Write(fill.Character);
            }
        }
    }

    public override void SetBufferContents(Coordinates origin, BufferCell[,] contents)
    {
        int rows = contents.GetLength(0);
        int cols = contents.GetLength(1);
        for (int y = 0; y < rows; y++)
        {
            for (int x = 0; x < cols; x++)
            {
                Console.SetCursorPosition(origin.X + x, origin.Y + y);
                Console.Write(contents[y, x].Character);
            }
        }
    }

    public override BufferCell[,] GetBufferContents(Rectangle rectangle)
    {
        int rows = rectangle.Bottom - rectangle.Top + 1;
        int cols = rectangle.Right - rectangle.Left + 1;
        var buffer = new BufferCell[rows, cols];

        for (int y = 0; y < rows; y++)
        {
            for (int x = 0; x < cols; x++)
            {
                buffer[y, x] = new BufferCell
                {
                    Character = ' ',
                    ForegroundColor = Console.ForegroundColor,
                    BackgroundColor = Console.BackgroundColor
                };
            }
        }

        return buffer;
    }

    public override void ScrollBufferContents(Rectangle source, Coordinates destination, Rectangle clip, BufferCell fill)
    {
        // Implement buffer scrolling logic
        throw new NotImplementedException();
    }
}