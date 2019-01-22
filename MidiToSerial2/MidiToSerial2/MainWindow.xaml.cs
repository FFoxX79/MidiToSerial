using System;
using System.Diagnostics;
using System.IO.Ports;
using System.Linq;
using System.Windows.Input;
using Midi;
using System.Collections.ObjectModel;
using System.Collections.Generic;

namespace MidiToSerial2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private readonly SerialPort _serialPort;

        private Midi.InputDevice _midiDevice;

        private ObservableCollection<string> _text = new ObservableCollection<string>();

        public MainWindow()
        {
            InitializeComponent();

            this.DataContext = _text;

            List<string> ports = SerialPort.GetPortNames().ToList();
            ports.Sort();

            _serialPort = _serialPort = new SerialPort(ports.Last(), 38400, Parity.None, 8, StopBits.One);
            _serialPort.Open();

            _midiDevice = Midi.InputDevice.InstalledDevices.First();

            _midiDevice.Open();

            _midiDevice.NoteOn += MidiDeviceOnNoteOn;
            _midiDevice.NoteOff += MidiDeviceOnNoteOff;
            _midiDevice.PitchBend += MidiDevicePitchBend;
            _midiDevice.ControlChange += MidiDeviceControlChange;

            _midiDevice.StartReceiving(null);
        }

        private void MidiDeviceControlChange(ControlChangeMessage msg)
        {
            byte[] buffer = { (byte)(128 + (3 << 4) + msg.Channel), (byte)msg.Control, (byte)msg.Value };
            _serialPort.Write(buffer, 0, 3);
        }

        private void MidiDevicePitchBend(PitchBendMessage msg)
        {
            byte[] buffer = { (byte)(128 + (6 << 4) + msg.Channel), (byte)(msg.Value & 0x7F), (byte)(msg.Value >> 7) };
            _serialPort.Write(buffer, 0, 3);
        }

        private void MidiDeviceOnNoteOff(NoteOffMessage msg)
        {
            byte[] buffer = { (byte)(128 + msg.Channel), (byte)(int)msg.Pitch, 127 };
            _serialPort.Write(buffer, 0, 3);
        }

        private void MidiDeviceOnNoteOn(NoteOnMessage msg)
        {
            byte[] buffer = new byte[] { (byte)(128 + (1 << 4) + msg.Channel), (byte)(int)msg.Pitch, (byte)msg.Velocity };
            _serialPort.Write(buffer, 0, 3);
        }

        private void MainWindow_OnKeyDown(object sender, KeyEventArgs e)
        {
            byte? note = KeyToNote(e.Key);
            byte? channel = KeyToChannel(e.Key);

            if (!e.IsRepeat)
            {
                if(note.HasValue && channel.HasValue)
                {
                    byte[] buffer = { (byte)(128 + (1 << 4) + channel.Value), note.Value, 127 };
                    _serialPort.Write(buffer, 0, 3);

                    Trace.WriteLine(string.Format("Note on - {0}", buffer[1]));
                }

                _text.Insert(0, string.Format("{0}: Key down - {1}", DateTime.Now.ToLongTimeString(), e.Key));
            }
        }

        private void MainWindow_OnKeyUp(object sender, KeyEventArgs e)
        {
            byte? note = KeyToNote(e.Key);
            byte? channel = KeyToChannel(e.Key);

            if (note.HasValue && channel.HasValue)
            {
                byte[] buffer = { (byte)(128 + channel.Value), note.Value, 127 };
                _serialPort.Write(buffer, 0, 3);
            }

            _text.Insert(0, string.Format("{0}: Key up - {1}", DateTime.Now.ToLongTimeString(), e.Key));
        }

        private void MainWindow_OnClosed(object sender, EventArgs e)
        {
            _midiDevice.StopReceiving();
            _midiDevice.Close();

            _serialPort.Close();
        }

        private static byte? KeyToNote(Key key)
        {
            switch (key)
            {
                case Key.Oem102:
                    return 35;
                case Key.Z:
                    return 36;
                case Key.S:
                    return 37;
                case Key.X:
                    return 38;
                case Key.D:
                    return 39;
                case Key.C:
                    return 40;
                case Key.V:
                    return 41;
                case Key.G:
                    return 42;
                case Key.B:
                    return 43;
                case Key.H:
                    return 44;
                case Key.N:
                    return 45;
                case Key.J:
                    return 46;
                case Key.M:
                    return 47;
                case Key.OemComma:
                    return 48;
                case Key.L:
                    return 49;
                case Key.OemPeriod:
                    return 50;
                case Key.OemSemicolon:
                    return 51;
                case Key.Oem2:
                    return 52;
                case Key.Q:
                    return 48;
                case Key.D2:
                    return 49;
                case Key.W:
                    return 50;
                case Key.D3:
                    return 51;
                case Key.E:
                    return 52;
                case Key.R:
                    return 53;
                case Key.D5:
                    return 54;
                case Key.T:
                    return 55;
                case Key.D6:
                    return 56;
                case Key.Y:
                    return 57;
                case Key.D7:
                    return 58;
                case Key.U:
                    return 59;
                case Key.I:
                    return 60;
                case Key.D9:
                    return 61;
                case Key.O:
                    return 62;
                case Key.D0:
                    return 63;
                case Key.P:
                    return 64;
                case Key.Oem4:
                    return 65;
                case Key.OemPlus:
                    return 66;
                case Key.Oem6:
                    return 67;
            }

            return null;
        }

        private static byte? KeyToChannel(Key key)
        {
            switch (key)
            {
                case Key.Oem102:
                    return 0;
                case Key.Z:
                    return 0;
                case Key.S:
                    return 0;
                case Key.X:
                    return 0;
                case Key.D:
                    return 0;
                case Key.C:
                    return 0;
                case Key.V:
                    return 0;
                case Key.G:
                    return 0;
                case Key.B:
                    return 0;
                case Key.H:
                    return 0;
                case Key.N:
                    return 0;
                case Key.J:
                    return 0;
                case Key.M:
                    return 0;
                case Key.OemComma:
                    return 0;
                case Key.L:
                    return 0;
                case Key.OemPeriod:
                    return 0;
                case Key.OemSemicolon:
                    return 0;
                case Key.Oem2:
                    return 0;
                case Key.Q:
                    return 1;
                case Key.D2:
                    return 1;
                case Key.W:
                    return 1;
                case Key.D3:
                    return 1;
                case Key.E:
                    return 1;
                case Key.R:
                    return 1;
                case Key.D5:
                    return 1;
                case Key.T:
                    return 1;
                case Key.D6:
                    return 1;
                case Key.Y:
                    return 1;
                case Key.D7:
                    return 1;
                case Key.U:
                    return 1;
                case Key.I:
                    return 1;
                case Key.D9:
                    return 1;
                case Key.O:
                    return 1;
                case Key.D0:
                    return 1;
                case Key.P:
                    return 1;
                case Key.Oem4:
                    return 1;
                case Key.OemPlus:
                    return 1;
                case Key.Oem6:
                    return 1;
            }

            return null;
        }
    }
}
