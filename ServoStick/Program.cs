using System;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.IO.Ports;
using System.Management;
using System.Threading;

namespace ServoStick
{
    class Program
    {
        public enum JoystickControlMode
        {
            J4WAY,
            J8WAY,
            J45DEGREES
        }

        static string AutodetectArduinoPort()
        {
            ManagementScope connectionScope = new ManagementScope();
            SelectQuery serialQuery = new SelectQuery("SELECT * FROM Win32_SerialPort");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(connectionScope, serialQuery);

            try
            {
                foreach (ManagementObject item in searcher.Get())
                {
                    string desc = item["Description"].ToString();
                    string deviceId = item["DeviceID"].ToString();

                    if (desc.Contains("Arduino"))
                    {
                        return deviceId;
                    }
                }
            }
            catch (ManagementException e)
            {
                /* Do Nothing */
            }

            return null;
        }

        static bool setJoystickMode(JoystickControlMode mode)
        {
            bool bSuccess = false;

            Console.WriteLine($"Writing joystick state {mode.ToString()}");

            String arduinoComPort = AutodetectArduinoPort();
            String angle = "";
            if (arduinoComPort != null)
            {
                Console.WriteLine($"Arduino Found on port: {arduinoComPort}");
                switch (mode)
                {
                    case JoystickControlMode.J4WAY:
                        angle = "065";
                        break;
                    case JoystickControlMode.J8WAY:
                        angle = "020";
                        break;
                    case JoystickControlMode.J45DEGREES:
                        angle = "020";
                        break;
                }

                SerialPort port = new SerialPort(arduinoComPort, 115200);
                port.ReadTimeout = 500;
                port.WriteTimeout = 500;

                port.Open();

                if (port.IsOpen)
                {
                    try
                    {
                        Console.WriteLine($"Connected to Arduino on {arduinoComPort}");
                        Console.WriteLine($"Setting Player 1 to {mode.ToString()}");
                        String p1Message = "1" + angle;
                        port.Write(p1Message);
                        Console.WriteLine($"Setting Player 2 to {mode.ToString()}");
                        Thread.Sleep(500);
                        String p2Message = "2" + angle;
                        port.Write(p1Message);
                    }
                    catch(Exception ex)
                    {
                        Console.WriteLine($"Error Writing Joystick States. {ex.Message}");
                    }

                    port.Close();
                }
                else
                {
                    Console.WriteLine($"Error opening Serial Port to Arduino on {arduinoComPort}");
                }

            }
            else{
                Console.WriteLine($"No Arduino Found");
            }

            return bSuccess;
        }

        static void Main(string[] args)
        {
            if (args.Length == 2)
            {
                String commandName = args[0];
                String argName = args[1].Trim().Replace(".zip", "");
                String controlsFile = "controls.json";

                switch (commandName)
                {
                    case "set":
                        if (argName.Trim() == "4")
                        {
                            setJoystickMode(JoystickControlMode.J4WAY);
                        }
                        else if (argName.Trim() == "45")
                        {
                            setJoystickMode(JoystickControlMode.J45DEGREES);
                        }
                        else if (argName.Trim() == "8")
                        {
                            setJoystickMode(JoystickControlMode.J8WAY);
                        }
                        else
                        {
                            Console.WriteLine($"Ignoring Joystick Set Commmand due to invalid joystick mode {argName}");
                        }
                        break;
                    case "game":
                        if (File.Exists(controlsFile))
                        {
                            try
                            {
                                JObject gameObj = null;

                                foreach (JObject obj in JObject.Parse(File.ReadAllText(controlsFile))["games"].Value<JArray>().Children<JObject>())
                                {
                                    if (obj["romname"].Value<String>() == argName)
                                    {
                                        gameObj = obj;
                                        break;
                                    }
                                }

                                object controlsString = gameObj["players"].Value<JArray>().Children().Select(p => p.Value<JToken>()).ToList()[0]["controls"].Children().Select(p => p.Value<JObject>()).ToList()[0].Value<JObject>()["name"].Value<String>();

                                switch (controlsString)
                                {
                                    case "4-way Joystick":
                                        setJoystickMode(JoystickControlMode.J4WAY);
                                        break;
                                    case "Diagonal 4-way Joystick":
                                        setJoystickMode(JoystickControlMode.J45DEGREES);
                                        break;
                                    case "8-way Joystick":
                                    default:
                                        setJoystickMode(JoystickControlMode.J8WAY);
                                        break;
                                }

                            }
                            catch
                            {
                                // Set 8 Way if we cant find game
                                Console.WriteLine($"Cant find game {argName}. Defaulting to 8-Way Joystick");
                                setJoystickMode(JoystickControlMode.J8WAY);
                            }
                        }
                        break;
                }
            }
            else
            {
                Console.WriteLine($"Usage: {args[0]} set (4|8|45) 45 = diagonal. {args[0]} name \"RomName\"");
            }
            
        }
    }
}
