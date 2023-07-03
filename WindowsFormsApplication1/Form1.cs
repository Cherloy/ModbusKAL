using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using Modbus.Device;
using Modbus;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        SerialPort sp;
        ModbusSerialMaster master;
        byte slaveID = 1;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            sp = new SerialPort(); //Create a new SerialPort object.
            master = ModbusSerialMaster.CreateRtu(sp);

            comboBox1.Text = "COM1";
            comboBox2.Text = "9600";
            sp.DataBits = 8;
            sp.Parity = Parity.None;
            sp.StopBits = StopBits.One;
            master.Transport.ReadTimeout = 300;
            master.Transport.WriteTimeout = 300;

            for(int i = 1; i <= 256; ++i)
            {
                comboBox1.Items.Add("COM" + i);
            } 
            button5.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int Parse_RD;
            ushort Reg_Nom_Rd = 0;
            ushort Reg_Num_Rd = 0;

            if (textBox1.Text != "" && textBox2.Text != "")
            {
                if (int.TryParse(textBox1.Text, out Parse_RD) && int.TryParse(textBox2.Text, out Parse_RD))
                {
                    Reg_Nom_Rd = Convert.ToUInt16(textBox1.Text);
                    Reg_Num_Rd = Convert.ToUInt16(textBox2.Text);
                    ModbusSerialMaster master = ModbusSerialMaster.CreateRtu(sp);

                    try
                    {
                        try
                        {
                            dataGridView1.Rows.Clear();
                        }
                        catch { }

                        ushort[] holding_register = master.ReadHoldingRegisters(slaveID, Reg_Nom_Rd, Reg_Num_Rd);
                        for (int i = 0; i < holding_register.Length; ++i)
                        {
                            dataGridView1.Rows.Add(Reg_Nom_Rd + i,holding_register[i]);  
                        }

                    }
                    catch (TimeoutException)
                    {
                        MessageBox.Show("Не обнаружено подчиненных устройств в выбранном порту\r\n Пожалуйста, выберите другой порт");
                    }
                    catch (SlaveException)
                    {
                        MessageBox.Show("Введено больше регистров, чем существует в подчиненном устройстве");
                    }
                    catch (ArgumentException)
                    {
                        MessageBox.Show("Введите ненулевое количество регистров");
                    }
                    catch (InvalidOperationException)
                    {
                        MessageBox.Show("Вы не подключились к порту");
                    }
                }
                else
                {
                    MessageBox.Show("Введите целочисленное значение");
                }
            }
            else
            {
                MessageBox.Show("Введите корректные данные");
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            int Parse_WR;
            ushort Reg_Nom_WR = 0;
            ushort Reg_Val_WR = 0;

            if (textBox3.Text != "" && textBox4.Text != "")
            {
                if (int.TryParse(textBox3.Text, out Parse_WR) && int.TryParse(textBox4.Text, out Parse_WR))
                {
                    Reg_Nom_WR = Convert.ToUInt16(textBox3.Text);
                    Reg_Val_WR = Convert.ToUInt16(textBox4.Text);
                    ModbusSerialMaster master = ModbusSerialMaster.CreateRtu(sp);

                    try
                    {
                        dataGridView1.Rows[Reg_Nom_WR].SetValues(Reg_Nom_WR, Reg_Val_WR);
                    }
                    catch{}

                    try
                    {
                        master.WriteSingleRegister(slaveID, Reg_Nom_WR, Reg_Val_WR);
                    }
                    catch (TimeoutException)
                    {
                        MessageBox.Show("Не обнаружено подчиненных устройств в выбранном порту\r\n Пожалуйста, выберите другой порт");
                    }
                    catch (SlaveException)
                    {
                        MessageBox.Show("Данного регистра не существует в подчиненном устройстве");
                    }
                    catch (InvalidOperationException)
                    {
                        MessageBox.Show("Вы не подключились к порту");
                    }
                    
                }
                else {
                MessageBox.Show("Введите одно целочисленное значение");  
                }
            }
            else
            {
                MessageBox.Show("Введите данные");
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Clear(); 
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            dataGridView1.Rows.Clear();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                try
                {
                    ModbusSerialMaster master = ModbusSerialMaster.CreateRtu(sp);
                    master.Dispose();
                }
                catch{}

                sp.PortName = comboBox1.Text;
                sp.BaudRate = Convert.ToUInt16(comboBox2.Text);
                sp.Open();
                label7.Text = "Есть соединение";
                label7.ForeColor = Color.Green;
                button5.Enabled = true;
            }
            catch
            {
                MessageBox.Show("Данный порт занят");
                label7.Text = "Подключение не удалось";
                label7.ForeColor = Color.Orange;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                ModbusSerialMaster master = ModbusSerialMaster.CreateRtu(sp);
                master.Dispose();
                label7.Text = "Нет соединения";
                label7.ForeColor = Color.Red;
                button5.Enabled = false;
            }
            catch
            {
                MessageBox.Show("Вы не были подключены к порту");
            }
        }

        private void Form1_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                master.Dispose();
            }
            catch {}
        }
    }
}
