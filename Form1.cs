using FluentModbus;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ParamsCopyer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // 初始化操作（如有）
        }

        /// <summary>
        /// 校验IP地址和端口
        /// </summary>
        private bool ValidateIpPort(string ip, int port)
        {
            string pattern = @"^(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}$";
            return Regex.IsMatch(ip, pattern) && port >= 1 && port <= 65535;
        }

        /// <summary>
        /// 读取寄存器并保存为XML
        /// </summary>
        private void button1_Click(object sender, EventArgs e)
        {
            string ip = textBox2.Text.Trim();
            int port = 502;
            byte slaveId = 1;
            ushort startAddress = 100;
            int totalCount = 300;
            int batchSize = 50;

            if (!ValidateIpPort(ip, port))
            {
                MessageBox.Show("目标IP或端口格式错误！", "错误");
                return;
            }

            progressBar1.Maximum = (int)Math.Ceiling((double)totalCount / batchSize);
            progressBar1.Value = 0;

            var client = new ModbusTcpClient();
            List<RegisterData> dataList = new List<RegisterData>();
            try
            {
                string connStr = $"{ip}:{port}";
                client.Connect(connStr, ModbusEndianness.BigEndian);

                for (int i = 0; i < totalCount; i += batchSize)
                {
                    ushort address = (ushort)(startAddress + i);
                    ushort readCount = (ushort)Math.Min(batchSize, totalCount - i);

                    if (readCount > 125) readCount = 125; // Modbus协议一次最多读取125个

                    var bytes = client.ReadHoldingRegisters(slaveId, address, readCount);

                    progressBar1.Value = (i / batchSize) + 1;
                    Application.DoEvents();
                    Thread.Sleep(500);

                    for (int j = 0; j < readCount; j++)
                    {
                        ushort value = (ushort)((bytes[2 * j] << 8) | bytes[2 * j + 1]);
                        dataList.Add(new RegisterData
                        {
                            address = address + j,
                            value = value
                        });
                    }
                }

                if (dataList.Count == 0)
                {
                    MessageBox.Show("未读取到任何数据！", "错误");
                    return;
                }

                // 保存XML
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Title = "保存配置文件";
                saveDialog.Filter = "XML文件|*.xml";
                saveDialog.FileName = "bmu_config.xml";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    var xml = new XElement("Registers",
                        dataList.Select(d => new XElement("Register",
                            new XAttribute("address", d.address),
                            new XAttribute("value", d.value)))
                    );
                    xml.Save(saveDialog.FileName);

                    MessageBox.Show("读取完成，已保存为 " + saveDialog.FileName, "成功");
                    progressBar1.Value = 0;
                }
                else
                {
                    progressBar1.Value = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("读取失败: " + ex.Message, "错误");
                LogError("读取异常", ex);
                progressBar1.Value = 0;
            }
            finally
            {
                client.Disconnect();
            }
        }

        /// <summary>
        /// 写入寄存器（仅写有变化的），数据来自XML
        /// </summary>
        private void button2_Click(object sender, EventArgs e)
        {
            string ip = textBox2.Text.Trim();
            int port = 502;
            byte slaveId = 1;

            if (!ValidateIpPort(ip, port))
            {
                MessageBox.Show("目标IP或端口格式错误！", "错误");
                return;
            }

            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Title = "选择配置文件";
            openDialog.Filter = "XML文件|*.xml";

            if (openDialog.ShowDialog() != DialogResult.OK)
                return;

            List<RegisterData> configList = new List<RegisterData>();
            try
            {
                var xml = XElement.Load(openDialog.FileName);
                foreach (var reg in xml.Elements("Register"))
                {
                    int addr;
                    ushort val;
                    if (!int.TryParse(reg.Attribute("address")?.Value, out addr) ||
                        !ushort.TryParse(reg.Attribute("value")?.Value, out val))
                    {
                        MessageBox.Show($"寄存器解析失败: 地址或数据非法！", "错误");
                        return;
                    }
                    if (addr < 0 || addr > 65535)
                    {
                        MessageBox.Show($"地址越界: {addr}", "错误");
                        return;
                    }
                    configList.Add(new RegisterData { address = addr, value = val });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("配置文件解析失败: " + ex.Message, "错误");
                LogError("XML解析异常", ex);
                return;
            }

            if (configList.Count == 0)
            {
                MessageBox.Show("没有可用的数据！", "错误");
                return;
            }

            // 去重、排序
            configList = configList
                .GroupBy(x => x.address)
                .Select(g => g.First())
                .OrderBy(x => x.address)
                .ToList();

            int batchReadSize = 60; // 目标设备最大安全读取数
            int totalCount = configList.Count;

            var client = new ModbusTcpClient();
            Dictionary<int, ushort> targetValues = new Dictionary<int, ushort>();
            try
            {
                string connStr = $"{ip}:{port}";
                client.Connect(connStr, ModbusEndianness.BigEndian);

                // 1. 读取目标设备当前寄存器值（批量读）
                var allAddresses = configList.Select(x => x.address).Distinct().OrderBy(a => a).ToList();
                int addrIdx = 0;
                while (addrIdx < allAddresses.Count)
                {
                    var batchAddrs = allAddresses.Skip(addrIdx).Take(batchReadSize).ToList();
                    ushort startAddr = (ushort)batchAddrs[0];
                    ushort count = (ushort)batchAddrs.Count;
                    var bytes = client.ReadHoldingRegisters(slaveId, startAddr, count);
                    for (int j = 0; j < count; j++)
                    {
                        ushort value = (ushort)((bytes[2 * j] << 8) | bytes[2 * j + 1]);
                        targetValues[startAddr + j] = value;
                    }
                    addrIdx += batchReadSize;
                    Thread.Sleep(500);
                }

                // 2. 找出需要写入的寄存器（值不同才写）
                var diffList = configList.Where(d =>
                    !targetValues.ContainsKey(d.address) || targetValues[d.address] != d.value).ToList();

                if (diffList.Count == 0)
                {
                    MessageBox.Show("目标设备参数已一致，无需写入。", "提示");
                    return;
                }

                // 3. 只写有变化的（逐条或小批量）
                int batchWriteSize = 8;
                progressBar2.Maximum = diffList.Count;
                progressBar2.Value = 0;

                int writtenCount = 0;
                int idx = 0;
                while (idx < diffList.Count)
                {
                    var batch = diffList.Skip(idx).Take(batchWriteSize).ToList();
                    // 检查批量地址连续性
                    bool isContinuous = true;
                    for (int k = 1; k < batch.Count; k++)
                    {
                        if (batch[k].address != batch[k - 1].address + 1)
                        {
                            isContinuous = false;
                            break;
                        }
                    }

                    if (!isContinuous)
                    {
                        // 不连续逐条写
                        foreach (var reg in batch)
                        {
                            client.WriteMultipleRegisters(slaveId, reg.address, new ushort[] { reg.value });
                            writtenCount++;
                            progressBar2.Value = writtenCount;
                            Application.DoEvents();
                            Thread.Sleep(500);
                        }
                    }
                    else
                    {
                        // 连续批量写
                        ushort[] values = batch.Select(x => x.value).ToArray();
                        ushort startAddress = (ushort)batch[0].address;
                        if (values.Length > 123)
                            values = values.Take(123).ToArray();
                        client.WriteMultipleRegisters(slaveId, startAddress, values);
                        writtenCount += values.Length;
                        progressBar2.Value = writtenCount;
                        Application.DoEvents();
                        Thread.Sleep(500);
                    }
                    idx += batchWriteSize;
                }

                MessageBox.Show($"写入完成，已更新 {diffList.Count} 项参数！", "成功");
                progressBar2.Value = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("写入失败: " + ex.Message, "错误");
                LogError("写入异常", ex);
                progressBar2.Value = 0;
            }
            finally
            {
                client.Disconnect();
            }
        }

        /// <summary>
        /// 错误日志记录到本地
        /// </summary>
        private void LogError(string title, Exception ex)
        {
            string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app_error.log");
            try
            {
                File.AppendAllText(logPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {title}: {ex.Message} {ex.StackTrace}{Environment.NewLine}");
            }
            catch { /* ignore log failure */ }
        }
    }

    public class RegisterData
    {
        public int address { get; set; }
        public ushort value { get; set; }
    }
}