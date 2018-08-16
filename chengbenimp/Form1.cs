using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOIExcelOpa;
using Sqlhelper;

namespace chengbenimp
{
    public partial class Form1 : Form
    {
        DataTable excelDt = null;
        string connStr = @"Data Source = u8testserver; Initial Catalog = UFDATA_003_2016; User ID = sa; Password=jiwu@2017";
        int count = 0;
        public Form1()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 检查数据
        /// </summary>
        private bool checkExcelData()
        {
            try
            {
                LogRichTextBox.logMesg("正在检查导入数据的格式。。。");
                
                    if (!excelDt.Columns.Contains("存货分类编码") || !excelDt.Columns.Contains("商品代码") || !excelDt.Columns.Contains("表面色") || !excelDt.Columns.Contains("金额"))
                    {
                        LogRichTextBox.logMesg("EXCEL文件中缺少存货分类编码、商品代码、表面色和金额等关键列！数据完整性检查失败！请核对要导入的EXCEL文件。", 1);
                        return false;
                    }
               
                LogRichTextBox.logMesg("数据完整,可以导入。");
            }
            catch (Exception e)
            {
                LogRichTextBox.logMesg("数据完整性检查程序出错！错误原因：" + e.ToString(), 2);
                return false;
            }
            return true;
        }
        public bool initDatabase()
        {
            ArrayList sqlstrList = new ArrayList();
            sqlstrList.Add(@"if exists(select name from sysobjects where name = N'IA_ProdAssTmpU8TESTSERVER1' And xtype = 'U') drop table [IA_ProdAssTmpU8TESTSERVER1]");
            sqlstrList.Add(@"select IA_ProdAss.* into IA_ProdAssTmpU8TESTSERVER1 from IA_ProdAss inner join rdrecord10 rdrecord on rdrecord.id=ia_prodass.id where isnull(cbaccounter,N'')=N'' and rdrecord.cvouchtype=N'10'");
            sqlstrList.Add(@"create index IA_index_IA_ProdAssTmpU8TESTSERVER1_ID on IA_ProdAssTmpU8TESTSERVER1 (autoid,id,cinvcode,cfree1,cfree2,cfree3,cfree4,cfree5,cfree6,cfree7,cfree8,cfree9,cfree10,iquantity,cbaccounter)");
            sqlstrList.Add(@"if exists(select name from sysobjects where name = N'IA_ProdAssTmpU8TESTSERVER1AA' And xtype = 'U') drop table [IA_ProdAssTmpU8TESTSERVER1AA]");
            sqlstrList.Add(@"select autoid,id,iunitcost,iquantity,iprice,iprice as icalprice,id as igroup  into IA_ProdAssTmpU8TESTSERVER1AA from IA_ProdAssTmpU8TESTSERVER1 where 1=0");
            sqlstrList.Add(@"create index IA_index_IA_ProdAssTmpU8TESTSERVER1AA_ID on IA_ProdAssTmpU8TESTSERVER1AA (autoid,id,igroup)");
            try
            {
                DbHelperSQL.ExecuteSqlTran(sqlstrList);
                return true;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 提交数据库
        /// </summary>
        public void commitDatabase()
        {
           
            ArrayList sqllist = new ArrayList();
            string startDate = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string endDate = dateTimePicker2.Value.ToString("yyyy-MM-dd");
            string jine = "0";
            string sqlstr = @"insert into IA_ProdAssTmpU8TESTSERVER1AA (autoid,id,iquantity,iunitcost,iprice,icalprice,igroup) select IA_ProdAss.autoid,IA_ProdAss.id,IA_ProdAss.iquantity,{金额},cast( {金额} * IA_ProdAss.iquantity as decimal(20,2)),{金额},{序号} from  (((RdRecord10 RdRecord LEFT JOIN IA_ProdAssTmpU8TESTSERVER1 IA_ProdAss ON RdRecord.ID=IA_ProdAss.ID)"
                            +@"LEFT JOIN Inventory ON IA_ProdAss.cInvCode=Inventory.cInvCode) "
                            +@"LEFT JOIN Warehouse ON RdRecord.cWhcode=Warehouse.cWhCode) " 
                            +@"left join ComputationUnit on inventory.cComunitCode= ComputationUnit.cComunitCode "
                            +@"Left join department on rdrecord.cdepcode = department.cdepcode " 
                            +@"Where(IA_ProdAss.cBAccounter= N''Or (IA_ProdAss.cBAccounter is null ))And RdRecord.cVouchType= N'10' "
                            +@"And RdRecord.dDate>= N'{当月开始日期}' And RdRecord.dDate<= N'{当月结束日期}' And IA_ProdAss.iQuantity>0 "
                            + @"And RdRecord.cWhCode in (N'2') AND (IA_ProdAss.iPrice=0 Or (IA_ProdAss.iPrice is null )) And IA_ProdAss.cInvCode= N'{存货分类编码}' " 
                            +@"and IA_ProdAss.cfree1 = N'{商品代码}' and IA_ProdAss.cfree2 = N'{表面色}' and ((IA_ProdAss.cfree3 is null) or IA_ProdAss.cfree3 = N'') "
                            +@"and ((IA_ProdAss.cfree4 is null) or IA_ProdAss.cfree4 = N'')  and ((IA_ProdAss.cfree5 is null) or IA_ProdAss.cfree5 = N'') "  
                            +@"and ((IA_ProdAss.cfree6 is null) or IA_ProdAss.cfree6 = N'')  and ((IA_ProdAss.cfree7 is null) or IA_ProdAss.cfree7 = N'') "
                            +@"and ((IA_ProdAss.cfree8 is null) or IA_ProdAss.cfree8 = N'')  and ((IA_ProdAss.cfree9 is null) or IA_ProdAss.cfree9 = N'') "  
                            +@"and ((IA_ProdAss.cfree10 is null) or IA_ProdAss.cfree10 = N'') ";
            string uptmpaa = @"update  IA_ProdAssTmpU8TESTSERVER1AA  set iprice = iprice +icalprice-isumprice from IA_ProdAssTmpU8TESTSERVER1AA aa inner join (select max(autoid) autoid,igroup,sum(iprice) isumprice from IA_ProdAssTmpU8TESTSERVER1AA group by igroup) bb on aa.autoid=bb.autoid";
            string uprd10sql = @"update rdrecords10 set iunitcost = aa.iunitcost,iprice=aa.iprice from rdrecords10 rdrecords inner join IA_ProdAssTmpU8TESTSERVER1AA aa on aa.autoid=rdrecords.autoid";
            int i = 0;
            count = 0;
            foreach (DataRow dr in excelDt.Rows)
            {
                i++;
                if (string.IsNullOrEmpty(dr["商品代码"].ToString()))
                {
                    continue;
                }
                else
                {
                    if (string.IsNullOrEmpty(dr["金额"].ToString()))
                        jine = "0";
                    else
                        jine = dr["金额"].ToString();
                    string sql = sqlstr.Replace("{金额}", jine).Replace("{序号}", i.ToString()).Replace("{当月开始日期}", startDate)
                                   .Replace("{当月结束日期}", endDate).Replace("{存货分类编码}", dr["存货分类编码"].ToString())
                                   .Replace("{表面色}", dr["表面色"].ToString()).Replace("{商品代码}", dr["商品代码"].ToString());
                    sqllist.Add(sql);
                    count++;
                }
            }
            sqllist.Add(uptmpaa);
            sqllist.Add(uprd10sql);
            DbHelperSQL.ExecuteSqlTran(sqllist);
        }
        private void openExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx|所有文件|*.*";
            ofd.ValidateNames = true;
            ofd.CheckPathExists = true;
            ofd.CheckFileExists = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string strFileName = ofd.FileName;
                try
                {
                    excelDt = NPOIExcel.readExcel(strFileName,true);
                    if (excelDt != null && excelDt.Rows.Count > 1)
                    {
                        LogRichTextBox.logMesg("文件『" + strFileName + "』读取成功！共读取【" + excelDt.Rows.Count.ToString() + "】条记录。");
                    }
                    else
                    {
                        LogRichTextBox.logMesg("文件『" + strFileName + "』读取失败！请检查EXCEL文件是否有数据。", 2);
                    }
                }
                catch(Exception ex)
                {
                    LogRichTextBox.logMesg("文件『" + strFileName + "』读取失败！失败原因为：" + ex.ToString(),2);
                    return;
                }                
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LogRichTextBox.richTextBoxRemote = this.richTextBox1;
            //this.ControlBox = false; //去所有按钮
            this.MaximizeBox = false;//去大最化按钮
            this.MinimizeBox = false;//去最小化按钮
            this.BackColor = Color.FromArgb(194, 216, 240);
            this.ShowIcon = false;

            checkBox1.Checked = false;
            dateTimePicker1.Value = (DateTime)datetimeCal.GetDateTimeStartByType("Month", DateTime.Now);
            dateTimePicker2.Value = (DateTime)datetimeCal.GetDateTimeEndByType("Month", DateTime.Now);

            Sqlhelper.DbHelperSQL.setConnectStr(connStr);
        }

        private void impU8_Click(object sender, EventArgs e)
        {
            if (!checkBox1.Checked)
            {
                LogRichTextBox.logMesg("请确认产成品成本分配的会计期间后，再导入！",1);
                return;
            }
            if (!checkExcelData())
                return;

            try
            {
                LogRichTextBox.logMesg("请等候，开始导入U8系统。。。");
                impU8.Enabled = false;
                if (initDatabase())
                    commitDatabase();
                else
                    LogRichTextBox.logMesg("初始化导入环境出错，请联系系统管理员！",2);
                LogRichTextBox.logMesg("数据导入成功！共导入【" + count.ToString() + "】条记录！");
                excelDt.Rows.Clear();//导入成功，清空数据
            }
            catch (Exception ex)
            {
                impU8.Enabled = true;
                LogRichTextBox.logMesg("程序出错！错误原因为：" + ex.ToString(), 2);
            }
        }

        private void 清除日志ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LogRichTextBox.LogClear();
        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            dateTimePicker1.Enabled = !checkBox1.Checked;
            dateTimePicker2.Enabled = !checkBox1.Checked;
        }
    }
}
