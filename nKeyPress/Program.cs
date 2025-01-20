namespace nKeyPress
{
    using System;
    using System.Threading;
    using System.Windows.Forms;
    using System.Runtime.InteropServices;
    using Microsoft.VisualBasic;

    public class NKeyPress
    {
        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private const byte PAGE_DOWN = 0x22;
        private const int KEYEVENTF_EXTENDEDKEY = 0x0001;
        private const int KEYEVENTF_KEYUP = 0x0002;

        private double initialRest = 5000;
        private double rest = 5000;
        private double minusMultiplier = 1.05;
        private double plusMultiplier = 0.95;
        private bool running = true;
        private int rounds = 100;
        private Thread workerThread;

        private enum KeyModifier
        {
            None = 0,
            Alt = 1,
            Control = 2,
            Shift = 4,
            WinKey = 8
        }

        private class HotKeyMessageFilter : IMessageFilter
        {
            private NKeyPress parent;
            public HotKeyMessageFilter(NKeyPress parent) { this.parent = parent; }

            public bool PreFilterMessage(ref Message m)
            {
                if (m.Msg == 0x0312)
                {
                    int id = m.WParam.ToInt32();
                    switch (id)
                    {
                        case 0: parent.AdjustRest(true); break;        // Ctrl-Minus
                        case 1: parent.AdjustRest(false); break;       // Ctrl-Plus
                        case 2: parent.StopLoop(); break;              // Ctrl-Shift-Minus
                        case 3: parent.SetPlusMultiplier(); break;     // Ctrl-Shift-Plus
                        case 4: parent.ResetRest(); break;             // Ctrl-Asterisk
                        case 5: parent.ExitProgram(); break;           // Ctrl-Slash
                        case 6: parent.SetMinusMultiplier(); break;    // Ctrl-Alt-Minus
                        case 7: parent.SetInitialRest(); break;        // Ctrl-Alt-Asterisk
                        case 8: parent.ResetVariables(); break;        // Ctrl-Shift-Asterisk
                        case 9: parent.SetRounds(); break;             // Ctrl-Alt-/
                    }
                    return true;
                }
                return false;
            }
        }

        public void Run()
        {
            RegisterHotKey(IntPtr.Zero, 0, (int)KeyModifier.Control, Keys.Subtract.GetHashCode());  // Ctrl-Minus
            RegisterHotKey(IntPtr.Zero, 1, (int)KeyModifier.Control, Keys.Add.GetHashCode());       // Ctrl-Plus
            RegisterHotKey(IntPtr.Zero, 2, (int)(KeyModifier.Control | KeyModifier.Shift), Keys.Subtract.GetHashCode()); // Ctrl-Shift-Minus
            RegisterHotKey(IntPtr.Zero, 3, (int)(KeyModifier.Control | KeyModifier.Shift), Keys.Add.GetHashCode());      // Ctrl-Shift-Plus
            RegisterHotKey(IntPtr.Zero, 4, (int)KeyModifier.Control, Keys.Multiply.GetHashCode());                      // Ctrl-Asterisk
            RegisterHotKey(IntPtr.Zero, 5, (int)KeyModifier.Control, Keys.Divide.GetHashCode());                        // Ctrl-Slash
            RegisterHotKey(IntPtr.Zero, 6, (int)(KeyModifier.Control | KeyModifier.Alt), Keys.Subtract.GetHashCode());  // Ctrl-Alt-Minus
            RegisterHotKey(IntPtr.Zero, 7, (int)(KeyModifier.Control | KeyModifier.Alt), Keys.Multiply.GetHashCode());  // Ctrl-Alt-Asterisk
            RegisterHotKey(IntPtr.Zero, 8, (int)(KeyModifier.Control | KeyModifier.Shift), Keys.Multiply.GetHashCode()); // Ctrl-Shift-Asterisk
            RegisterHotKey(IntPtr.Zero, 9, (int)(KeyModifier.Control | KeyModifier.Alt), Keys.Divide.GetHashCode());    // Ctrl-Alt-/ 

            Application.AddMessageFilter(new HotKeyMessageFilter(this));
            Application.Run();
        }

        private void StartLoop()
        {
            if (workerThread == null || !workerThread.IsAlive)
            {
                running = true;
                workerThread = new Thread(() =>
                {
                    int i = 0;
                    while (running)
                    {
                        Thread.Sleep((int)rest);
                        if (!running) break;
                        SendPageDown();
                        Console.WriteLine($"Page Down sent. Rest: {rest / 1000} seconds. Round: {i++}");

                        if (i == rounds)
                        {
                            StopLoop();
                        }
                    }
                    Console.WriteLine("Loop stopped.");
                    workerThread = null;
                });
                workerThread.Start();
                Console.WriteLine("Loop started.");
            }
            else
            {
                Console.WriteLine("Loop is already running.");
            }
        }

        private void StopLoop()
        {
            running = false;
            Console.WriteLine("Stopping loop...");
            if (workerThread != null && workerThread.IsAlive)
            {
                workerThread.Join();
            }
        }

        private void SendPageDown()
        {
            keybd_event(PAGE_DOWN, 0x4F, KEYEVENTF_EXTENDEDKEY | 0, UIntPtr.Zero);
            keybd_event(PAGE_DOWN, 0x4F, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, UIntPtr.Zero);
        }

        public void AdjustRest(bool increase)
        {
            if (increase) rest *= minusMultiplier;
            else rest *= plusMultiplier;
            Console.WriteLine($"Rest adjusted to: {rest / 1000} seconds.");
        }

        public void ResetRest()
        {
            rest = initialRest;
            Console.WriteLine($"Rest reset to: {rest / 1000} seconds.");
        }

        public void SetMinusMultiplier()
        {
            string input = Interaction.InputBox("Please enter the new minus multiplier:", "Set Minus Multiplier", minusMultiplier.ToString());

            if (!string.IsNullOrEmpty(input))
            {
                if (double.TryParse(input, out double newMultiplier))
                {
                    if (newMultiplier > 0)
                    {
                        minusMultiplier = newMultiplier;
                        Console.WriteLine($"Minus multiplier set to: {minusMultiplier}");
                    }
                    else
                    {
                        MessageBox.Show("Multiplier must be greater than zero.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Invalid input. Please enter a valid number.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public void SetPlusMultiplier()
        {
            string input = Interaction.InputBox("Please enter the new plus multiplier:", "Set Plus Multiplier", plusMultiplier.ToString());

            if (!string.IsNullOrEmpty(input))
            {
                if (double.TryParse(input, out double newMultiplier))
                {
                    if (newMultiplier > 0)
                    {
                        plusMultiplier = newMultiplier;
                        Console.WriteLine($"Plus multiplier set to: {plusMultiplier}");
                    }
                    else
                    {
                        MessageBox.Show("Multiplier must be greater than zero.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Invalid input. Please enter a valid number.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public void SetInitialRest()
        {
            string input = Interaction.InputBox("Please enter the new initial rest (in milliseconds):", "Set Initial Rest", initialRest.ToString());

            if (!string.IsNullOrEmpty(input))
            {
                if (int.TryParse(input, out int newInitialRest))
                {
                    if (newInitialRest > 0)
                    {
                        initialRest = newInitialRest;
                        rest = initialRest; // Also reset the current rest value
                        Console.WriteLine($"Initial rest set to: {initialRest} milliseconds.");
                    }
                    else
                    {
                        MessageBox.Show("Initial rest must be greater than zero.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Invalid input. Please enter a valid integer.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void SetRounds()
        {
            string input = Interaction.InputBox("Please enter the new rounds:", "Set Rounds", rounds.ToString());

            if (!string.IsNullOrEmpty(input))
            {
                if (int.TryParse(input, out int newRounds))
                {
                    if (newRounds > 0)
                    {
                        rounds = newRounds;
                        Console.WriteLine($"Rounds set to: {rounds}");
                    }
                    else
                    {
                        MessageBox.Show("Rounds must be greater than zero.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Invalid input. Please enter a valid integer.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public void ResetVariables()
        {
            initialRest = 5000;
            rest = 5000;
            minusMultiplier = 1.05;
            plusMultiplier = 0.95;
            rounds = 100;
            Console.WriteLine("Variables reset to initial state.");
        }

        public void ExitProgram()
        {
            StopLoop();
            UnregisterHotKey(IntPtr.Zero, 0);
            UnregisterHotKey(IntPtr.Zero, 1);
            UnregisterHotKey(IntPtr.Zero, 2);
            UnregisterHotKey(IntPtr.Zero, 3);
            UnregisterHotKey(IntPtr.Zero, 4);
            UnregisterHotKey(IntPtr.Zero, 5);
            UnregisterHotKey(IntPtr.Zero, 6);
            UnregisterHotKey(IntPtr.Zero, 7);
            UnregisterHotKey(IntPtr.Zero, 8);
            UnregisterHotKey(IntPtr.Zero, 9);
            Console.WriteLine("Exiting program.");
            Application.Exit();
        }

        public static void Main(string[] args)
        {
            NKeyPress kp = new NKeyPress();
            kp.StartLoop();
            kp.Run();
        }
    }
}