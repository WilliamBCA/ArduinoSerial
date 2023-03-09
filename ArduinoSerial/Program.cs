namespace ArduinoSerial;
using System.IO.Ports;
using System.Threading;
class Program
{
    static void Main(string[] args)
    {
        var arduinoListener = new ArduinoListener();
        var thread = new Thread(arduinoListener.Listen);
        thread.Start();
        Console.WriteLine("Press Esc to exit...");
        while (Console.ReadKey().Key != ConsoleKey.Escape) { }
        arduinoListener.Close();
    }
}

public class ArduinoListener
{
    private readonly SerialPort _serialPort;
    private readonly SheetsHelper _sheetsHelper;

    public ArduinoListener()
    {
        _serialPort = new SerialPort("/dev/cu.usbserial-0001", 115200); // Replace with your Arduino's serial port and baud rate
        _sheetsHelper = new SheetsHelper("1pkPWcTLgkyEXoCOtn9cmniUPybBvQ71YM6WGOcYpDC4");
    }

    public void Listen()
    {
        _serialPort.DataReceived += SerialCallback;
        _serialPort.Open();

        while (true)
        {
            // Wait for a short time to avoid blocking the thread
            Thread.Sleep(10);
        }
    }

    private void SerialCallback(object sender, SerialDataReceivedEventArgs e)
    {
        string line = _serialPort.ReadLine();
        var splitLine = line.Split(';');
        if (splitLine.Length < 3) return;
        _sheetsHelper.AddRowToSheetAsync(splitLine);
    }

    public void Close()
    {
        _serialPort.Close();
    }
}

