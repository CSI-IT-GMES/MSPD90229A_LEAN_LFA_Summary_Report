using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Xml;

using JPlatform.Client.Library6.interFace;
using JPlatform.Client;
using JPlatform.Client.Controls6;
using JPlatform.Client.JBaseForm6;
using JPlatform.Client.CSIGMESBaseform6;
using DevExpress.XtraReports.UI;
using DevExpress.XtraGrid.Columns;
using System.Globalization;
using DevExpress.XtraPrinting.Control;
using DevExpress.XtraPrinting;
using System.Reflection;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.Utils;
using System.Data.OleDb;
using System.IO;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraCharts;
using System.Diagnostics;
namespace CSI.GMES.PD
{
    public partial class MSPD90229A_POP : CSIGMESBaseform6 
    {
        #region Variable

        public bool _firstLoad = true, _isSaved = false, _null_mline = false;
        public DataTable _dtFty = null;

        #endregion

        public MSPD90229A_POP()
        {
            InitializeComponent();
        }

        #region Load Data
        protected override void OnLoad(EventArgs e)
        {
            try
            {
                base.OnLoad(e);
                _firstLoad = true;

                cboMonth.EditValue = DateTime.Now.ToString();
                LoadDataCbo(cboFactory, "Factory", "Q_FTY");
                LoadDataCbo(cboPlant, "Plant", "Q_LINE");
                LoadDataCbo(cboLine, "Line", "Q_MLINE");

                Check_Confirm_Status();

                _firstLoad = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Check_Confirm_Status()
        {
            try
            {
                DataTable dt = GetData("Q_STATUS");
                chkAlert.Checked = false;

                if (dt != null && dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["CONFIRM_YN"].ToString().Equals("Y"))
                    {
                        chkAlert.Checked = true;
                    }
                }
            }
            catch {}
        }

        public bool Check_Alllow_Update()
        {
            try
            {
                bool _result = false;

                DataTable dt = GetData("Q_POPUP_CHANGE");

                if (dt != null && dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["ALLOW_UPDATE"].ToString().Equals("Y"))
                    {
                        _result = true;
                    }
                }

                return _result;
            }
            catch {
                return false;
            }
        }

        private void LoadDataCbo(LookUpEditEx argCbo, string _cbo_nm, string _type)
        {
            try
            {
                DataTable dt = GetData(_type);
                if (dt == null)
                {
                    argCbo.Properties.Columns.Clear();
                    argCbo.Properties.DataSource = dt;

                    if (_type.Equals("Q_MLINE"))
                    {
                        lbLine.Visible = false;
                        cboLine.Visible = false;
                        _null_mline = true;
                    }

                    return;
                }

                if (_type.Equals("Q_MLINE"))
                {
                    lbLine.Visible = true;
                    cboLine.Visible = true;
                    _null_mline = false;
                }
                else if (_type.Equals("Q_FTY"))
                {
                    _dtFty = dt.Copy();
                }

                string columnCode = dt.Columns[0].ColumnName;
                string columnName = dt.Columns[1].ColumnName;
                string captionCode = "Code";
                string captionName = _cbo_nm;

                argCbo.Properties.Columns.Clear();
                argCbo.Properties.DataSource = dt;
                argCbo.Properties.ValueMember = columnCode;
                argCbo.Properties.DisplayMember = columnName;
                argCbo.Properties.Columns.Add(new DevExpress.XtraEditors.Controls.LookUpColumnInfo(columnCode));
                argCbo.Properties.Columns.Add(new DevExpress.XtraEditors.Controls.LookUpColumnInfo(columnName));
                argCbo.Properties.Columns[columnCode].Visible = false;
                argCbo.Properties.Columns[columnCode].Width = 10;
                argCbo.Properties.Columns[columnCode].Caption = captionCode;
                argCbo.Properties.Columns[columnName].Caption = captionName;
                argCbo.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }

        private DataTable GetData(string argType)
        {
            try
            {
                P_MSPD90229A_Q proc = new P_MSPD90229A_Q();
                DataTable dtData = null;

                if (argType.Equals("Q_PERMISS"))
                {
                    string _userID = SessionInfo.UserID;
                    dtData = proc.SetParamData(dtData, argType, _userID, "", "", "");
                }
                else
                {
                    string _factory = cboFactory.EditValue == null ? "" : cboFactory.EditValue.ToString();
                    string _plant = cboPlant.EditValue == null ? "" : cboPlant.EditValue.ToString();
                    string _line = _null_mline ? "000" : cboLine.EditValue == null ? "" : cboLine.EditValue.ToString();

                    dtData = proc.SetParamData(dtData, argType, _factory, _plant, _line, cboMonth.yyyymmdd);
                }

                ResultSet rs = CommonCallQuery(dtData, proc.ProcName, proc.GetParamInfo(), false, 90000, "", true);
                if (rs == null || rs.ResultDataSet == null || rs.ResultDataSet.Tables.Count == 0 || rs.ResultDataSet.Tables[0].Rows.Count == 0)
                {
                    return null;
                }
                return rs.ResultDataSet.Tables[0];
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return null;
            }
        }

        public void SetBrowserMain(JPlatform.Client.Library6.interFace.IBrowserMain JbrowserMain)
        {
            this._browserMain = JbrowserMain;
        }

        public bool CheckIsSaved()
        {
            return _isSaved;
        }

        #endregion

        #region Events

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dlr;

                if (_dtFty == null || _dtFty.Rows.Count < 1) return;

                DataTable _dtPermiss = GetData("Q_PERMISS");
                if (_dtPermiss == null || _dtPermiss.Rows.Count < 1)
                {
                    MessageBox.Show("Bạn không có quyền thực hiện chức năng này!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                bool _allow_update = Check_Alllow_Update();
                if (!_allow_update)
                {
                    MessageBox.Show("Bạn chỉ được thay đổi dữ liệu của:\n- Tháng hiện tại. \n- Tháng trước nếu ngày hiện tại không lớn hơn ngày 5", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                dlr = MessageBox.Show("Bạn có muốn Save không?", "Save", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dlr == DialogResult.Yes)
                {
                    bool result = SaveData("Q_UNCONFIRM");
                    if (result)
                    {
                        MessageBoxW("Save successfully!", IconType.Information);
                        _isSaved = true;
                    }
                    else
                    {
                        MessageBoxW("Save failed!", IconType.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public bool SaveData(string _type)
        {
            try
            {
                bool _result = true;
                DataTable dtData = null;
                P_MSPD90229A_S proc = new P_MSPD90229A_S();
                string machineName = $"{SessionInfo.UserName}|{ Environment.MachineName}|{GetIPAddress()}";
                string _plant_cd = "";

                JPlatform.Client.CSIGMESBaseform6.frmSplashScreenWait frmSplash = new JPlatform.Client.CSIGMESBaseform6.frmSplashScreenWait();
                frmSplash.Show();

                for (int iRow = 0; iRow < _dtFty.Rows.Count; iRow++)
                {
                    if (_dtFty.Rows[iRow]["CODE"].ToString().Equals(cboFactory.EditValue.ToString()))
                    {
                        _plant_cd = _dtFty.Rows[iRow]["PLANT_CD"].ToString();
                        break;
                    }
                }

                dtData = proc.SetParamData(dtData,
                          _type,
                          _plant_cd,
                          cboPlant.EditValue.ToString(),
                          cboLine.EditValue.ToString(),
                          cboMonth.yyyymm,
                          "",
                          "",
                          "",
                          "",
                          null,
                          "",
                          machineName,
                          "CSI.GMES.PD.MSPD90229A_POP");
                _result = CommonProcessSave(dtData, proc.ProcName, proc.GetParamInfo(), null);

                frmSplash.Close();
                return _result;
            }
            catch (Exception ex)
            {
                pbProgressHide();
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        private void cboFactory_EditValueChanged(object sender, EventArgs e)
        {
            if (!_firstLoad)
            {
                LoadDataCbo(cboPlant, "Plant", "Q_LINE");
            }
        }

        private void cboPlant_EditValueChanged(object sender, EventArgs e)
        {
            if (!_firstLoad)
            {
                LoadDataCbo(cboLine, "Line", "Q_MLINE");
            }
        }

        private void cboLine_EditValueChanged(object sender, EventArgs e)
        {
            if (!_firstLoad)
            {
                Check_Confirm_Status();
            }
        }

        private void cboMonth_EditValueChanged(object sender, EventArgs e)
        {
            if (!_firstLoad)
            {
                Check_Confirm_Status();
            }
        }

        #endregion

        #region Database

        public class P_MSPD90229A_Q : BaseProcClass
        {
            public P_MSPD90229A_Q()
            {
                // Modify Code : Procedure Name
                _ProcName = "LMES.P_MSPD90229A_Q";
                ParamAdd();
            }
            private void ParamAdd()
            {
                _ParamInfo.Add(new ParamInfo("@ARG_WORK_TYPE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_FTY", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_LINE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_MLINE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_YMD", "Varchar", 100, "Input", typeof(System.String)));
            }
            public DataTable SetParamData(DataTable dataTable,
                                        System.String ARG_WORK_TYPE,
                                        System.String ARG_FTY,
                                        System.String ARG_LINE,
                                        System.String ARG_MLINE,
                                        System.String ARG_YMD)
            {
                if (dataTable == null)
                {
                    dataTable = new DataTable(_ProcName);
                    foreach (ParamInfo pi in _ParamInfo)
                    {
                        dataTable.Columns.Add(pi.ParamName, pi.TypeClass);
                    }
                }
                // Modify Code : Procedure Parameter
                object[] objData = new object[] {
                                                ARG_WORK_TYPE,
                                                ARG_FTY,
                                                ARG_LINE,
                                                ARG_MLINE,
                                                ARG_YMD
                };
                dataTable.Rows.Add(objData);
                return dataTable;
            }
        }

        public class P_MSPD90229A_S : BaseProcClass
        {
            public P_MSPD90229A_S()
            {
                // Modify Code : Procedure Name
                _ProcName = "LMES.P_MSPD90229A_S";
                ParamAdd();
            }
            private void ParamAdd()
            {
                _ParamInfo.Add(new ParamInfo("@ARG_TYPE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_PLANT", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_LINE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_MLINE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_YMD", "Varchar2", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_GROUP", "Varchar2", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_ITEM", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_MAX_SCORE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_RESULT", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_PHOTO", "BLOB", 900000, "Input", typeof(byte[])));
                _ParamInfo.Add(new ParamInfo("@ARG_COUNTER", "Varchar2", 0, "Input", typeof(System.String)));

                _ParamInfo.Add(new ParamInfo("@ARG_CREATE_PC", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@ARG_CREATE_PROGRAM_ID", "Varchar", 100, "Input", typeof(System.String)));
            }
            public DataTable SetParamData(DataTable dataTable,
                                        System.String ARG_TYPE,
                                        System.String ARG_PLANT,
                                        System.String ARG_LINE,
                                        System.String ARG_MLINE,
                                        System.String ARG_YMD,
                                        System.String ARG_GROUP,
                                        System.String ARG_ITEM,
                                        System.String ARG_MAX_SCORE,
                                        System.String ARG_RESULT,
                                        byte[] ARG_PHOTO,
                                        System.String ARG_COUNTER,
                                        System.String ARG_CREATE_PC,
                                        System.String ARG_CREATE_PROGRAM_ID)
            {
                if (dataTable == null)
                {
                    dataTable = new DataTable(_ProcName);
                    foreach (ParamInfo pi in _ParamInfo)
                    {
                        dataTable.Columns.Add(pi.ParamName, pi.TypeClass);
                    }
                }
                // Modify Code : Procedure Parameter
                object[] objData = new object[] {
                    ARG_TYPE,
                    ARG_PLANT,
                    ARG_LINE,
                    ARG_MLINE,
                    ARG_YMD,
                    ARG_GROUP,
                    ARG_ITEM,
                    ARG_MAX_SCORE,
                    ARG_RESULT,
                    ARG_PHOTO,
                    ARG_COUNTER,
                    ARG_CREATE_PC,
                    ARG_CREATE_PROGRAM_ID
                };
                dataTable.Rows.Add(objData);
                return dataTable;
            }
        }

        #endregion
    }
}
