using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Data.SqlClient;

namespace QuestionnaireTool {
    public partial class frmMain : Form {
        string SurveyMonth = "";
        string Connnstring = "";
        IWorkbook workbook = null;
        FileStream fileStream = null;
        DataTable table = new DataTable();
        //
        bool AddToDase = true;
        bool AddToTxt = false;
        bool AddToPhone = false;
        public frmMain() {
            InitializeComponent();
        }
        private void frmMain_Load(object sender, EventArgs e) {
            //id,Phone,type,SurveyMonth,City,AddTime,Dimension
            //特别注意：列的次序要和目标数据表一致，否则会导致导入数据失败或者数据错位。
            table.Columns.Add("id", typeof(System.Guid));
            table.Columns.Add("Phone", typeof(string));
            table.Columns.Add("type", typeof(int));
            table.Columns.Add("SurveyMonth", typeof(string));
            table.Columns.Add("City", typeof(string));
            table.Columns.Add("AddTime", typeof(DateTime));
            table.Columns.Add("Dimension", typeof(string));//维度


            textBox2.AppendText("手机来源格式：\r\n");
            textBox2.AppendText("\t\t四个Sheet,名称规定如：4G新增,4G故障,家宽新增,家宽故障 \t(不必4个齐全，必须有一个Sheet)\r\n");
            textBox2.AppendText("\t\t每个Sheet,第一列必须是手机，第2列必须是市县名称，手机列不能空，不能是其他字符，市县列可以空\r\n");
            //
            textBox4.Text = System.DateTime.Now.ToString("yyyyMM");
            textBox5.Text = System.Configuration.ConfigurationManager.ConnectionStrings["Con_SatisfySurvey_181"].ConnectionString;
            //

        }
        private void button1_Click(object sender, EventArgs e) {
            if (backgroundWorker1.IsBusy) {
                MessageBox.Show("上次的导入还未完成");
                return;
            }
            if (backgroundWorker2.IsBusy) {
                MessageBox.Show("正在发布中，请等待……");
                return;
            }
            if (workbook == null) {
                MessageBox.Show("请先导入文件");
                return;
            }
            SurveyMonth = textBox4.Text.Trim();
            Connnstring = textBox5.Text.Trim();

            if (SurveyMonth == "") {
                MessageBox.Show("调查月份不能空");
                return;
            }
            if (SurveyMonth == "201808") {
                MessageBox.Show("调查月份不能是初始月份201808");
                return;
            }
            if (!IsYearMonth(SurveyMonth)) {
                MessageBox.Show("调查月份不是正确的年份月份格式");
                return;
            }

            if (Connnstring == "") {
                MessageBox.Show("数据库连接字符串不能空");
                return;
            }
            AddToDase = checkBox1.Checked;
            AddToTxt = checkBox2.Checked;
            AddToPhone = checkBox4.Checked;
            //
            if (!AddToDase && !AddToTxt && !AddToPhone) {
                MessageBox.Show("【导入到数据库】、【输出SQL脚本】和【输出手机到单独文件】3项必须选择一个");
                return;
            }
            if (AddToTxt || AddToPhone) {
                if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + "\\脚本")) {
                    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + "\\脚本");
                }
            }
            if (AddToDase) {
                //检测数据库连接是否可用
                if (!CheckConnIsOK()) {
                    MessageBox.Show("连接数据库失败,请检查连接字符串");
                    return;
                }
                //检测PhoneConfig表是否已经存在本月手机
                //检测MainSubject表是否已经存在本月题目
                //检测RewardPool表是否已经存在本月
            }
            textBox3.Clear();
            textBox3.AppendText("正在导入……\r\n");
            textBox3.AppendText("正在处理数据……\r\n");
            pictureBox1.Visible = true;
            backgroundWorker2.RunWorkerAsync();
        }

        private void button2_Click(object sender, EventArgs e) {

            if (backgroundWorker1.IsBusy) {
                MessageBox.Show("上次的excel加载还未完成");
                return;
            }
            if (backgroundWorker2.IsBusy) {
                MessageBox.Show("上次的excel导出还未完成");
                return;
            }
            if (backgroundWorker3.IsBusy) {
                MessageBox.Show("上次的复制还未完成");
                return;
            }
            {
                if (fileStream != null) {
                    fileStream.Close();
                    fileStream = null;
                }
                if (workbook != null) {
                    workbook.Close();
                    workbook = null;
                }
                //
                var dlg = new OpenFileDialog();
                dlg.CheckFileExists = true;
                dlg.DefaultExt = "*.xlsx";
                dlg.CheckPathExists = true;
                dlg.Filter = "2007版(*.ex1)|*.xlsx|2003版(*.xls)|*.xls|所有文件(*.*)|*.*";
                dlg.Multiselect = false;
                if (dlg.ShowDialog() == DialogResult.OK) {
                    textBox2.Clear();
                    textBox2.AppendText("正在加载Excel，文件越大越慢，请耐心等待……");
                    pictureBox1.Visible = true;
                    textBox1.Text = dlg.FileName;
                    backgroundWorker1.RunWorkerAsync(dlg.FileName);
                }
            }
        }



        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e) {
            var fileName = e.Argument.ToString();
            fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            if (fileName.IndexOf(".xlsx") > 0) {
                //2007版
                workbook = new XSSFWorkbook(fileStream);
            } else if (fileName.IndexOf(".xls") > 0) {
                //2003版
                workbook = new HSSFWorkbook(fileStream);
            }
            StringBuilder sb = new StringBuilder();
            int sheetcnt = workbook.NumberOfSheets;
            int cnt4G = 0;
            int cntJK = 0;
            for (int i = 0; i < sheetcnt; i++) {
                var sheet = workbook.GetSheetAt(i);
                var shtName = sheet.SheetName.Trim().Replace("'", "").Replace("-", "").Replace(" ", "").ToLower();
                var dctype = (shtName.Contains("4g") || shtName.Contains("无线")) ? 2 : 1;

                var lieshu = "";
                if (sheet.LastRowNum > 0) {
                    lieshu = $"列数：{ sheet.GetRow(0).LastCellNum }，";
                    if (dctype == 2) {
                        cnt4G += sheet.LastRowNum;
                    } else {
                        cntJK += sheet.LastRowNum;
                    }
                }
                sb.AppendLine($"第{(i + 1)}个Sheet：{sheet.SheetName}，{lieshu}行数：{sheet.LastRowNum}");
                if (sheet.LastRowNum > 0) {
                    var row = sheet.GetRow(0);
                    if (row.LastCellNum >= 0) {
                        var phone = row.GetCell(0).ToString().Trim();//第一列 手机号码
                        var city = row.LastCellNum > 0 && row.GetCell(1) != null ? row.GetCell(1).ToString().Trim() : ""; //第一列 市县
                        if (!IsHandset(phone)) {
                            sb.AppendLine($"\t\t\t ****警告**** 第一列数据不是手机号码：" + phone);
                        }
                    } else {
                        sb.AppendLine($"\t\t\t ***警告**** 缺少列数，第1列必须是手机，第2列必须是市县");
                    }
                } else {
                    sb.AppendLine($"\t\t\t ***警告**** 行数为0");
                }
            }
            //
            //sb.AppendLine($"\t");
            sb.AppendLine($"\t4G 手机数量：{cnt4G}");
            sb.AppendLine($"\t家宽 手机数量：{cntJK}");
            //
            e.Result = sb.ToString();
            fileStream.Close();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            pictureBox1.Visible = false;
            if (e.Error == null) {
                textBox2.Clear();
                if (e.Result != null) {
                    textBox2.AppendText(e.Result.ToString());
                }
            } else {
                MessageBox.Show(e.Error.Message);
            }
        }

        /// <summary>
        /// backgroundWorker2
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e) {
            int sheetcnt = workbook.NumberOfSheets;
            //
            HashSet<string> hasPhone = new HashSet<string>();
            HashSet<string> hasKipPhone = new HashSet<string>();
            var addtime = DateTime.Now;
            // 
            int repeatCnt = 0;
            int oldImprtCnt = 0;
            StringBuilder sb = new StringBuilder();
            StringBuilder sbPhone = new StringBuilder();
            StringBuilder kipPhone = new StringBuilder();
            var sqlfname = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            //Excel最大行数
            int maxExcelRow = 1048576;
            int exRowTotal = 0;
            int partNo = 1;
            //
            var savepath = System.Windows.Forms.Application.StartupPath;
            var error = "";
            int phones = 0;

            bool newImport = checkBox3.Checked;//true 全新导入，false 补充导入
            //
            DataTable dtphones = new DataTable();
            //
            if (AddToDase) {
                phones = GetPhoneCount(this.SurveyMonth, out error);
                if (phones < 0) {
                    new Exception(error);
                    return;
                } else if (phones > 0) {
                    SetLog(this.SurveyMonth + $" 数据库中手机 已经存在(数量{phones})");
                    if (newImport) {
                        var dlgret = MessageBox.Show(this.SurveyMonth + $" 数据库中手机 已经存在(数量{phones})，\r\n按[确定] 将清除现有数据后重新导入，\r\n按[取消] 将取消操作", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                        if (dlgret == DialogResult.Cancel) {
                            SetLog("已经取消导入手机");
                            return;
                        }
                        //删除旧的
                        int delret = DeletePhones(this.SurveyMonth, out error);
                        if (delret < 0) {
                            new Exception(error);
                            return;
                        } else if (delret > 0) {
                            //删除旧的
                            SetLog("已删除" + this.SurveyMonth + " 旧记录数量：" + delret);
                        }
                    } else {
                        //需要判断现在补加的和之前的是否重复
                        dtphones = GetPhoneSet(this.SurveyMonth, out error);
                    }
                }
            }
            int kiprow = 0;
            for (int k = 0; k < sheetcnt; k++) {
                var sheet = workbook.GetSheetAt(k);
                var shtName = sheet.SheetName.Trim().Replace("'", "").Replace("-", "").Replace(" ", "").ToLower();
                SetLog($"正在处理第{(k + 1)}个Sheet：{shtName}，共{sheet.LastRowNum}行……");
                //
                table.Rows.Clear();
                sb.Clear();
                sbPhone.Clear();
                int txtrow = 0;
                int phonerow = 0;
                for (int i = 0; i < sheet.LastRowNum; i++) {
                    var row = sheet.GetRow(i);
                    if (row != null) {
                        if (row.LastCellNum >= 0) {
                            var phone = row.GetCell(0).ToString().Trim();//第一列 手机号码
                            var city = row.LastCellNum > 0 && row.GetCell(1) != null ? row.GetCell(1).ToString().Trim() : ""; //第一列 市县
                            if (IsHandset(phone)) {
                                if (hasPhone.Add(phone)) {
                                    var dctype = (shtName.Contains("4g") || shtName.Contains("无线")) ? 2 : 1;
                                    // 
                                    bool oldImp = false;//数据库中是否已经存在手机号码，
                                    if (!newImport) {
                                        //补充导入,需要判断现在补加的和之前的是否重复
                                        if (dtphones.Select($"phone='{phone}'").Length > 0) {
                                            oldImprtCnt++;
                                            oldImp = true;
                                        }
                                    }
                                    //
                                    if (!oldImp && AddToDase) {
                                        var trow = table.NewRow();
                                        trow["id"] = System.Guid.NewGuid().ToString();
                                        trow["Phone"] = phone;      //手机号码
                                        trow["City"] = city;        //市县
                                        trow["Dimension"] = shtName;// 维度
                                        trow["type"] = dctype;
                                        trow["SurveyMonth"] = this.SurveyMonth;
                                        trow["AddTime"] = addtime;
                                        //
                                        table.Rows.Add(trow);
                                    }
                                    if (!oldImp && AddToTxt) {
                                        var sql = $@"insert into [PhoneConfig] ([Id],[Phone],[City],[Dimension],[type],[SurveyMonth],[AddTime]) values ('{System.Guid.NewGuid()}','{phone}', '{city}','{shtName}',{dctype},'{this.SurveyMonth}','{addtime}');";
                                        sb.AppendLine(sql);
                                        txtrow++;
                                        if (txtrow > 1000) {
                                            File.AppendAllText(savepath + "\\脚本\\" + sqlfname + "_" + shtName + ".txt", sb.ToString(), Encoding.Default);
                                            sb.Clear();
                                            txtrow = 0;
                                        }
                                    }
                                    if (!oldImp && AddToPhone) {
                                        var sql = $@"{phone},{city},{shtName}";
                                        sbPhone.AppendLine(sql);
                                        phonerow++;
                                        exRowTotal++;
                                        if (exRowTotal >= maxExcelRow) {

                                            //
                                            File.AppendAllText(savepath + "\\脚本\\" + sqlfname + "_手机_P" + partNo + ".csv", sbPhone.ToString(), Encoding.Default);
                                            sbPhone.Clear();
                                            phonerow = 0;
                                            //
                                            partNo++;
                                            exRowTotal = 0;
                                        } else {
                                            if (phonerow > 1000) {
                                                File.AppendAllText(savepath + "\\脚本\\" + sqlfname + "_手机_P" + partNo + ".csv", sbPhone.ToString(), Encoding.Default);
                                                sbPhone.Clear();
                                                phonerow = 0;
                                            }
                                        }
                                    }
                                } else {
                                    //重复
                                    if (AddToPhone) {
                                        if (hasKipPhone.Add(phone)) {
                                            repeatCnt++;
                                            kiprow++;
                                            var sql = $@"{phone},{city},{shtName},重复";
                                            kipPhone.AppendLine(sql);
                                            if (kiprow > 1000) {
                                                File.AppendAllText(savepath + "\\脚本\\" + sqlfname + "_手机(重复).csv", kipPhone.ToString(), Encoding.Default);
                                                kipPhone.Clear();
                                                kiprow = 0;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                }//for 
                if (AddToDase) {
                    SetLog($"正在导入数据库，共{table.Rows.Count}行……");
                    table.TableName = "PhoneConfig";
                    if (table.Rows.Count > 0) {
                        Insert(table);
                        SetLog("导入数据库成功");
                    } else {
                        SetLog("空表忽略");
                    }
                }
                if (AddToTxt) {
                    if (txtrow > 0) {
                        File.AppendAllText(savepath + "\\脚本\\" + sqlfname + "_" + shtName + ".txt", sb.ToString(), Encoding.Default);
                        //
                        sb.Clear();
                        txtrow = 0;
                    }
                }
                if (AddToPhone) {

                    if (phonerow > 0) {
                        File.AppendAllText(savepath + "\\脚本\\" + sqlfname + "_手机_P" + partNo + ".csv", sbPhone.ToString(), Encoding.Default);
                        //
                        sbPhone.Clear();
                        phonerow = 0;
                    }
                    //
                    if (kiprow > 0) {
                        File.AppendAllText(savepath + "\\脚本\\" + sqlfname + "_手机(重复).csv", kipPhone.ToString(), Encoding.Default);
                        //
                        kipPhone.Clear();
                        kiprow = 0;
                    }
                }


            }//for


            //
            SetLog("");
            SetLog("过滤【数据表】重复数量：" + oldImprtCnt);
            SetLog("过滤【源文件】重复数量：" + repeatCnt);
            //
            if (AddToDase) {
                SetLog("导入数据库成功数量：" + (hasPhone.Count - oldImprtCnt));
            }
            if (AddToTxt || AddToPhone) {
                SetLog("脚本保存位置：" + savepath + "\\脚本");
            }
            if (AddToTxt) {
                SetLog("导出脚本成功数量：" + (hasPhone.Count - oldImprtCnt));
            }
            if (AddToPhone) {
                SetLog("导出手机成功数量：" + (hasPhone.Count - oldImprtCnt));
            }
            SetLog("");
            if (AddToDase) {
                phones = GetPhoneCount(this.SurveyMonth, out error);
                if (phones > 0) {
                    SetLog(this.SurveyMonth + " 数据库手机总数量：" + phones);
                }
            }
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            pictureBox1.Visible = false;
            if (e.Error == null) {
                if (e.Result != null) {
                    textBox3.AppendText(e.Result.ToString());
                }
            } else {
                MessageBox.Show(e.Error.Message);
            }
        }

        /// <summary>
        /// 将 <see cref="DataTable"/> 的数据批量插入到数据库中。
        /// </summary>
        /// <param name="dataTable">要批量插入的 <see cref="DataTable"/>。</param>
        /// <param name="batchSize">每批次写入的数据量。</param>
        public void Insert(DataTable dataTable, int batchSize = 10000) {
            if (this.table.Rows.Count == 0) {
                return;
            }
            using (var connection = new SqlConnection(this.Connnstring)) {
                SqlTransaction tran = null;
                try {
                    connection.Open();
                    tran = connection.BeginTransaction();
                    using (var bulk = new SqlBulkCopy(connection, SqlBulkCopyOptions.Default, tran) {
                        DestinationTableName = dataTable.TableName,
                        BatchSize = batchSize
                    }) {
                        bulk.WriteToServer(dataTable);
                        bulk.Close();
                    }
                    tran.Commit();
                } catch (Exception exp) {
                    if (tran != null)
                        tran.Rollback();
                    Exception ex = exp;
                    while (ex.InnerException != null) {
                        ex = ex.InnerException;
                    }
                    throw new Exception(exp.Message + "，" + ex.InnerException);
                } finally {
                    connection.Close();
                }
            }
        }

        private void SetLog(string log) {
            this.Invoke(new Action(() => {
                textBox3.AppendText($"{System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")} {log}\r\n");
            }));
        }

        /// <summary>
        /// 判断是否手机号码
        /// </summary>
        /// <param name="str_handset"></param>
        /// <returns></returns>
        private bool IsHandset(string handset) {

            return System.Text.RegularExpressions.Regex.IsMatch(handset, @"^[1]+[3,4,5,6,7,8,9]+\d{9}");

        }
        private bool IsYearMonth(string handset) {

            if (handset.Length != 6)
                return false;
            string inputTime = handset.Substring(0, 4) + "-" + handset.Substring(4, 2) + "-01";
            DateTime dateTime = new DateTime();
            return DateTime.TryParse(inputTime, out dateTime);

        }

        private bool CheckConnIsOK() {
            bool isok = false;
            using (var connection = new SqlConnection(this.Connnstring)) {
                try {
                    connection.Open();
                    if (connection.State == ConnectionState.Open) {
                        isok = true;
                    }
                } catch (Exception exp) {
                    MessageBox.Show("连接失败：" + exp.Message);
                } finally {
                    connection.Close();
                }
            }
            return isok;
        }
        private int GetPhoneCount(string month, out string error) {
            error = "";
            int count1 = 0;
            using (var connection = new SqlConnection(this.Connnstring)) {
                try {
                    connection.Open();
                    var sql1 = "select count(*) FROM [PhoneConfig] where [SurveyMonth]='" + this.SurveyMonth + "'";
                    SqlCommand cmd = new SqlCommand(sql1, connection);
                    var obj1 = cmd.ExecuteScalar();
                    if (obj1 != null) {
                        count1 = Convert.ToInt32(obj1);
                    }
                } catch (Exception exp) {
                    error = "连接失败：" + exp.Message;
                    count1 = -1;
                } finally {
                    connection.Close();
                }
            }
            return count1;
        }
        private DataTable GetPhoneSet(string month, out string error) {
            error = "";
            DataTable dt = new DataTable();
            using (var connection = new SqlConnection(this.Connnstring)) {
                try {
                    connection.Open();
                    var sql1 = "select [Phone] FROM [PhoneConfig] where [SurveyMonth]='" + this.SurveyMonth + "'";
                    SqlCommand cmd = new SqlCommand(sql1, connection);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                    da.Dispose();

                } catch (Exception exp) {
                    error = "连接失败：" + exp.Message;
                } finally {
                    connection.Close();
                }
            }
            return dt;
        }
        private int DeletePhones(string month, out string error) {
            error = "";
            int count1 = 0;
            using (var connection = new SqlConnection(this.Connnstring)) {
                try {
                    connection.Open();
                    var sql1 = "delete FROM [PhoneConfig] where [SurveyMonth]='" + this.SurveyMonth + "'";
                    SqlCommand cmd = new SqlCommand(sql1, connection);
                    var obj1 = cmd.ExecuteNonQuery();
                    count1 = obj1;
                } catch (Exception exp) {
                    error = "删除" + this.SurveyMonth + "旧记录异常：" + exp.Message;
                    count1 = -1;
                } finally {
                    connection.Close();
                }
            }
            return count1;
        }
        private void button3_Click(object sender, EventArgs e) {
            if (backgroundWorker1.IsBusy) {
                MessageBox.Show("上次的导入还未完成");
                return;
            }
            if (backgroundWorker2.IsBusy) {
                MessageBox.Show("正在发布中，请等待……");
                return;
            }
            if (backgroundWorker3.IsBusy) {
                MessageBox.Show("正在复制中，请等待……");
                return;
            }
            SurveyMonth = textBox4.Text.Trim();
            Connnstring = textBox5.Text.Trim();

            if (SurveyMonth == "") {
                MessageBox.Show("调查月份不能空");
                return;
            }
            if (SurveyMonth == "201808") {
                MessageBox.Show("调查月份不能是初始月份201808");
                return;
            }
            if (!IsYearMonth(SurveyMonth)) {
                MessageBox.Show("调查月份不是正确的年份月份格式");
                return;
            }

            if (Connnstring == "") {
                MessageBox.Show("数据库连接字符串不能空");
                return;
            }
            pictureBox1.Visible = true;
            backgroundWorker3.RunWorkerAsync();
        }
        /// <summary>
        /// backgroundWorker3
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void backgroundWorker3_DoWork(object sender, DoWorkEventArgs e) {
            Hashtable hastb = new Hashtable();
            hastb["errmsg"] = "";
            hastb["retadd1"] = 0;
            hastb["retadd2"] = 0;
            using (var connection = new SqlConnection(this.Connnstring)) {
                try {
                    connection.Open();
                    if (connection.State != ConnectionState.Open) {
                        hastb["errmsg"] = "连接失败,请重试";
                        e.Result = hastb;
                        return;
                    }
                    //检测MainSubject表是否已经存在本月题目
                    //检测RewardPool表是否已经存在本月
                    var sql1 = "select count(*) FROM [MainSubject] where [SurveyMonth]='" + this.SurveyMonth + "'";
                    var sql2 = "select count(*) FROM [RewardPool] where [SurveyMonth]='" + this.SurveyMonth + "'";
                    SqlCommand cmd = new SqlCommand(sql1, connection);
                    var obj1 = cmd.ExecuteScalar();
                    int count1 = 0;
                    if (obj1 != null) {
                        count1 = Convert.ToInt32(obj1);
                    }
                    cmd = new SqlCommand(sql2, connection);
                    var obj2 = cmd.ExecuteScalar();
                    int count2 = 0;
                    if (obj2 != null) {
                        count2 = Convert.ToInt32(obj2);
                    }
                    if (count1 > 0) {
                        SetLog("");
                        SetLog(this.SurveyMonth + $" 的题目 已经存在(数量{count1})");

                        var dlgret = MessageBox.Show(this.SurveyMonth + $" 的题目 已经存在(数量{count1})，\r\n按[确定] 将清除现有数据后重新生成，\r\n按[取消] 将取消操作", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                        if (dlgret == DialogResult.Cancel) {
                            hastb["errmsg"] = "cancel";
                            e.Result = hastb;
                            return;
                        }
                        //删除旧的题目
                        var sqldel1 = $"delete FROM [MainSubject] where [SurveyMonth]='{this.SurveyMonth}'";
                        cmd = new SqlCommand(sqldel1, connection);
                        var retdel1 = cmd.ExecuteNonQuery();
                        SetLog(this.SurveyMonth + $" 的题目 删除完成(删除行数{retdel1})");
                        if (retdel1 == 0) {
                            hastb["errmsg"] = $"删除{this.SurveyMonth}的旧题目 失败，请重试";
                            e.Result = hastb;
                            return;
                        }
                    }
                    if (count2 > 0) {
                        SetLog(this.SurveyMonth + $" 的奖池 已经存在(数量{count2})");
                        var sqldel2 = $"delete FROM [RewardPool] where [SurveyMonth]='{this.SurveyMonth}'";
                        cmd = new SqlCommand(sqldel2, connection);
                        var retdel2 = cmd.ExecuteNonQuery();
                        SetLog(this.SurveyMonth + $" 的奖池 删除完成(删除行数{retdel2})");
                        if (retdel2 == 0) {
                            hastb["errmsg"] = $"删除{this.SurveyMonth}的旧奖池 失败，请重试";
                            e.Result = hastb;
                            return;
                        }
                    }
                    var addtime = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    var titles = "[Id],[SurveyType],[Title],[AnswerType],[IsDisable],[IsDisableUnPrize],[Dimension]";
                    //题目
                    var sqladd1 = $@" insert into [MainSubject] ({titles},[SurveyMonth],[AddTime])
                                      select {titles},'{this.SurveyMonth}' as [SurveyMonth],'{addtime}' as [AddTime] from [MainSubject] where [IsDisable]=0 and [SurveyMonth]='201808'
                                    ";
                    //奖池
                    var sqladd2 = $@" insert into [RewardPool] ([Id],[Reward],[Token],[SurveyMonth],[AddTime])
                                      select NewID() as id,[Reward],0 as [Token],'{this.SurveyMonth}' as [SurveyMonth],'{addtime}' as [AddTime] from [RewardPool] where [SurveyMonth]='201808'
                                    ";
                    cmd = new SqlCommand(sqladd1, connection);
                    var retadd1 = cmd.ExecuteNonQuery();

                    SetLog("");
                    SetLog(this.SurveyMonth + $" 的题目 复制完成(复制行数{retadd1})");
                    cmd = new SqlCommand(sqladd2, connection);
                    var retadd2 = cmd.ExecuteNonQuery();
                    SetLog(this.SurveyMonth + $" 的奖池 复制完成(复制行数{retadd2})");
                    //
                    hastb["retadd1"] = retadd1;
                    hastb["retadd2"] = retadd2;

                } catch (Exception exp) {
                    hastb["errmsg"] = "出现异常：" + exp.Message;
                } finally {
                    connection.Close();
                }
            }
            e.Result = hastb;
        }

        private void backgroundWorker3_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            pictureBox1.Visible = false;
            if (e.Error == null) {
                Hashtable hastb = (Hashtable)e.Result;
                var errmsg = hastb["errmsg"].ToString();
                if (errmsg != "") {
                    if (errmsg != "cancel")
                        MessageBox.Show(this, errmsg);
                } else {
                    var retadd1 = Convert.ToInt32(hastb["retadd1"]);
                    var retadd2 = Convert.ToInt32(hastb["retadd2"]);
                    if (retadd1 > 0 && retadd2 > 0) {
                        MessageBox.Show(this, this.SurveyMonth + $" 的题目与奖池复制成功");
                    } else {
                        if (retadd1 == 0) {
                            MessageBox.Show(this, this.SurveyMonth + $" 的题目 复制失败，请重试");
                        }
                        if (retadd2 == 0) {
                            MessageBox.Show(this, this.SurveyMonth + $" 的奖池 复制失败，请重试");
                        }
                    }
                }
            } else {
                MessageBox.Show(this, e.Error.Message);
            }

        }
        private void button6_Click(object sender, EventArgs e) {
            var pt = System.Windows.Forms.Application.StartupPath + "\\脚本";
            System.Diagnostics.Process.Start("explorer.exe", pt);
        }
        private void button4_Click(object sender, EventArgs e) {

            SurveyMonth = textBox4.Text.Trim();
            Connnstring = textBox5.Text.Trim();

            if (SurveyMonth == "") {
                MessageBox.Show("调查月份不能空");
                return;
            }
            if (!IsYearMonth(SurveyMonth)) {
                MessageBox.Show("调查月份不是正确的年份月份格式");
                return;
            }

            if (Connnstring == "") {
                MessageBox.Show("数据库连接字符串不能空");
                return;
            }

            string error = "";
            int phones = GetPhoneCount(this.SurveyMonth, out error);
            if (error == "") {
                textBox3.AppendText(this.SurveyMonth + " 月手机数量为 " + phones + "\r\n");
            } else {
                MessageBox.Show(error);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e) {
            checkBox3.Visible = checkBox1.Checked;
        }

        private void textBox4_TextChanged(object sender, EventArgs e) {
            textBox6.Text = textBox4.Text;
        }

        private void button5_Click(object sender, EventArgs e) {
            if (backgroundWorker1.IsBusy) {
                MessageBox.Show("上次的导入还未完成");
                return;
            }
            if (backgroundWorker2.IsBusy) {
                MessageBox.Show("正在发布中，请等待……");
                return;
            }
            if (backgroundWorker3.IsBusy) {
                MessageBox.Show("正在复制中，请等待……");
                return;
            }
            SurveyMonth = textBox4.Text.Trim();
            Connnstring = textBox5.Text.Trim();

            if (SurveyMonth == "") {
                MessageBox.Show("调查月份不能空");
                return;
            }
            //if (SurveyMonth == "201808") {
            //    MessageBox.Show("调查月份不能是初始月份201808");
            //    return;
            //}
            if (!IsYearMonth(SurveyMonth)) {
                MessageBox.Show("调查月份不是正确的年份月份格式");
                return;
            }

            if (Connnstring == "") {
                MessageBox.Show("数据库连接字符串不能空");
                return;
            }

            var mobile = textBox7.Text.Trim();
            if (mobile == "") {
                MessageBox.Show("手机号码不能空");
                return;
            }
            //
            DeleteResult(mobile);
        }

        private void DeleteResult(string mobile) {
            using (var connection = new SqlConnection(this.Connnstring)) {
                try {
                    connection.Open();
                    if (connection.State != ConnectionState.Open) {
                        MessageBox.Show("连接失败,请重试");
                        return;
                    }
                    //检测MainSubject表是否已经存在本月题目
                    //检测RewardPool表是否已经存在本月
                    var sql = "select count(id) as cnt from [SurveyResult] where  [SurveyMonth]='" + this.SurveyMonth + "' and phone='" + mobile + "'";
                    SqlCommand cmd = new SqlCommand(sql, connection);
                    var obj1 = cmd.ExecuteScalar();
                    int count1 = 0;
                    if (obj1 != null) {
                        count1 = Convert.ToInt32(obj1);
                    }
                    if (count1 > 0) {
                        var dlgret = MessageBox.Show(this.SurveyMonth + $" 月 {mobile}  已经提交了 {count1}份问卷记录，\r\n按[确定] 将清除现有问卷记录，\r\n按[取消] 将取消操作", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
                        if (dlgret == DialogResult.Cancel) {
                            return;
                        }
                        DataTable dt = new DataTable();
                        var sql0 = "select id,[Reward] from [SurveyResult] where  [SurveyMonth]='" + this.SurveyMonth + "' and phone='" + mobile + "'";
                        SqlCommand cmd1 = new SqlCommand(sql0, connection);
                        SqlDataAdapter da = new SqlDataAdapter(cmd1);
                        da.Fill(dt);
                        int r_cnt = 0;
                        int d_cnt = 0;
                        int o_cnt = 0;
                        int c_cnt = 0;
                        //遍历主记录表，规则是只有一个手机一个月份只有一条记录
                        for (int i = 0; i < dt.Rows.Count; i++) {
                            var SurveyResultId = dt.Rows[i]["id"].ToString();
                            var dbReward = dt.Rows[i]["Reward"].ToString();
                            //
                            var sql1 = "select id from [SurveyDetail] where SurveyResultId='" + SurveyResultId + "' and [SurveyMonth]='" + this.SurveyMonth + "'";
                            cmd1.CommandText = sql1;
                            DataTable dt2 = new DataTable();
                            da = new SqlDataAdapter(cmd1);
                            da.Fill(dt2);

                            //1、遍历 SurveyDetail 表，删除SurveyOption明细
                            for (int i2 = 0; i2 < dt2.Rows.Count; i2++) {
                                var SurveyDetailId = dt2.Rows[i2]["id"].ToString();
                                //删除SurveyOption明细
                                var sql2 = "delete from [SurveyOption] where SurveyDetailId='" + SurveyDetailId + "' and [SurveyMonth]='" + this.SurveyMonth + "'";
                                cmd1.CommandText = sql2;
                                o_cnt += cmd1.ExecuteNonQuery();
                            }
                            //2、删除 SurveyDetail
                            var sql3 = "delete from [SurveyDetail] where SurveyResultId='" + SurveyResultId + "' and [SurveyMonth]='" + this.SurveyMonth + "'";
                            cmd1.CommandText = sql3;
                            d_cnt += cmd1.ExecuteNonQuery();
                            //
                            //
                            //3、删除主记录
                            var sql4 = "delete from [SurveyResult] where id='" + SurveyResultId + "' and [SurveyMonth]='" + this.SurveyMonth + "' and phone='" + mobile + "'";
                            cmd1.CommandText = sql4;
                            r_cnt += cmd1.ExecuteNonQuery();
                            //恢复奖池记录[RewardPool],规则确定只能恢复一条记录
                            var sql5 = "update [RewardPool] set Token=0 where id in (select top 1 id from [RewardPool] where  [Reward]=" + dbReward + " and [Token]=1 and [SurveyMonth]='" + this.SurveyMonth + "') and [Token]=1 and [SurveyMonth]='" + this.SurveyMonth + "'";
                            cmd1.CommandText = sql5;
                            c_cnt = cmd1.ExecuteNonQuery();
                        }
                        var msg = $"删除完成：\r\n删除主表:{r_cnt},\r\n删除详单:{d_cnt},\r\n删除明细:{o_cnt},\r\n恢复奖池记录:{c_cnt}";
                        MessageBox.Show(this.SurveyMonth + " 月" + mobile + "的问卷记录 " + msg);
                    } else {
                        MessageBox.Show(this.SurveyMonth + " 月" + mobile + "未发现问卷调查记录");
                    }
                } catch (Exception exp) {
                    MessageBox.Show("出现异常：" + exp.Message);
                } finally {
                    connection.Close();
                }
            }
        }


    }
}
