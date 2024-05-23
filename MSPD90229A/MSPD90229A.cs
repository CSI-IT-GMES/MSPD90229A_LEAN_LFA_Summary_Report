using DevExpress.Data;
using DevExpress.Utils;
using DevExpress.XtraCharts;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using JPlatform.Client.Controls6;
using JPlatform.Client.CSIGMESBaseform6;
using JPlatform.Client.JBaseForm6;
using JPlatform.Client.Library6.interFace;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace CSI.GMES.PD
{
    public partial class MSPD90229A : CSIGMESBaseform6//JERPBaseForm
    {
        public bool _firstLoad = true, _dateLoad = false;
        public MyCellMergeHelper _Helper = null;
        bool _allow_confirm = false, _null_mline = false, _is_other = false;
        public int _tab = 0;
        public OpenFileDialog openFileDialog = new OpenFileDialog();
        public string _last_1_month = "", _last_2_month = "", _format_cd = "";
        public string base64String = "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";

        public MSPD90229A()
        {
            InitializeComponent();
        }

        protected override void OnLoad(EventArgs e)
        {
            _firstLoad = true;

            base.OnLoad(e);
            NewButton = false;
            AddButton = false;
            DeleteRowButton = false;
            SaveButton = true;
            DeleteButton = false;
            PreviewButton = false;
            PrintButton = false;

            cboDate.EditValue = DateTime.Now.ToString();
            cboMonth.EditValue = DateTime.Now.ToString();
            panTop.BackColor = Color.FromArgb(240, 240, 240);
            tabControl.BackColor = Color.FromArgb(240, 240, 240);

            InitCombobox();
            FormatLayout();

            _firstLoad = false;
        }

        public void FormatLayout()
        {
            if (_tab.Equals(0))
            {
                panTop.Height = 50;

                cboMonth.Visible = true;
                cboDate.Visible = false;

                cboMonth.Location = new Point(68, 13);
                cboDate.Location = new Point(551, 13);

                lbFactory.Visible = false;
                lbPlant.Visible = false;
                lbLine.Visible = false;
                btnExport.Visible = true;

                cboFactory.Visible = false;
                cboPlant.Visible = false;
                cboLine.Visible = false;

                lbDate.Text = "Month";
                btnExport.Location = new Point(410, 13);
                lbDate.Location = new Point(3, 13);
                cboDate.Location = new Point(68, 13);

                lbGreen.Location = new Point(540, 13);
                lbYellow.Location = new Point(702, 13);
                lbRed.Location = new Point(849, 13);
                lbBlack.Location = new Point(996, 13);

                lbGreen.Text = "90% <= Rate <= 100%";
                lbYellow.Text = "80% <= Rate < 90%";
                lbRed.Text = "70% <= Rate < 80%";
                lbBlack.Text = "< 70%";

                lbGroup.Visible = true;
                cboGroup.Visible = true;

                lbGroup.Location = new Point(200, 13);
                cboGroup.Location = new Point(265, 13);

                btnConfirm.Visible = false;
                btnConfirm.Location = new Point(888, 8);
            }
            else if (_tab.Equals(1))
            {
                panTop.Height = 75;

                cboMonth.Visible = false;
                cboDate.Visible = true;

                cboDate.Location = new Point(68, 8);
                cboMonth.Location = new Point(551, 8);

                lbFactory.Visible = true;
                lbPlant.Visible = true;
                lbLine.Visible = true;
                btnExport.Visible = false;

                cboFactory.Visible = true;
                cboPlant.Visible = true;
                cboLine.Visible = true;

                lbDate.Text = "Date";
                lbDate.Location = new Point(3, 8);
                lbFactory.Location = new Point(206, 8);
                lbPlant.Location = new Point(3, 40);
                lbLine.Location = new Point(206, 40);

                cboDate.Location = new Point(68, 8);
                cboFactory.Location = new Point(271, 8);
                cboPlant.Location = new Point(68, 40);
                cboLine.Location = new Point(271, 40);

                lbGreen.Location = new Point(420, 40);
                lbYellow.Location = new Point(582, 40);
                lbRed.Location = new Point(729, 40);
                lbBlack.Location = new Point(876, 40);

                lbGroup.Visible = false;
                cboGroup.Visible = false;

                lbGroup.Location = new Point(687, 8);
                cboGroup.Location = new Point(752, 8);

                btnConfirm.Visible = true;
                btnConfirm.Location = new Point(420, 8);

                if (_null_mline)
                {
                    lbLine.Visible = false;
                    cboLine.Visible = false;

                    lbGreen.Location = new Point(220, 40);
                    lbYellow.Location = new Point(382, 40);
                    lbRed.Location = new Point(529, 40);
                    lbBlack.Location = new Point(676, 40);
                }
            }
        }

        #region [Start Button Event Code By UIBuilder]

        public override void QueryClick()
        {
            try
            {
                pbProgressShow();
                _allow_confirm = false;
                _is_other = false;
                _last_1_month = "";
                _last_2_month = "";
                _format_cd = "";

                if (_tab.Equals(0))
                {
                    InitControls(grdSummary);
                    DataTable _dtSource = GetData("Q_SUMMARY");
                    DataTable _dtChart = GetData("Q_SUMMARY_CHART");

                    if (_dtChart != null && _dtChart.Rows.Count > 0)
                    {
                        fn_load_chart(_dtChart);
                    }
                    else
                    {
                        chartData.DataSource = null;
                        while (chartData.Series[0].Points.Count > 0)
                        {
                            chartData.Series[0].Points.Clear();
                        }
                    }

                    if (_dtSource != null && _dtSource.Rows.Count > 0)
                    {
                        var distinctValues = _dtSource.AsEnumerable()
                                .Select(row => new
                                {
                                    GRP_NM = row.Field<string>("GRP_NM"),
                                    LINE_CD = row.Field<string>("LINE_CD"),
                                    LINE_NM = row.Field<string>("LINE_NM"),
                                })
                                .Distinct();
                        DataTable _dtHead = LINQResultToDataTable(distinctValues).Select("", "").CopyToDataTable();
                        CreateDetailGrid(grdSummary, gvwSummary, _dtHead);
                        DataTable _dtf = Binding_Data(_dtSource);
                        SetData(grdSummary, _dtf);
                        Formart_Grid_Summary();
                    }
                }
                else if (_tab.Equals(1))
                {
                    InitControls(grdDetail);
                    btnConfirm.Enabled = false;
                    DataTable _dtSource = GetData("Q_DETAIL");
                    DataTable _dtCheck = GetData("Q_ALLOW");
                    DataTable _dtMonth = GetData("Q_SMONTH");
                    DataTable _dtFormat = GetData("Q_FORMAT");

                    /////// Get Checklist Format
                    if (_dtFormat != null && _dtFormat.Rows.Count > 0)
                    {
                        _format_cd = _dtFormat.Rows[0]["FORMAT_CD"].ToString();
                        if (_format_cd.Equals("F003"))
                        {
                            _is_other = true;
                            lbGreen.Text = "90% <= Rate <= 100%";
                            lbYellow.Text = "85% <= Rate < 90%";
                            lbRed.Text = "80% <= Rate < 85%";
                            lbBlack.Text = "< 80%";
                        }
                        else
                        {
                            _is_other = false;
                            lbGreen.Text = "90% <= Rate <= 100%";
                            lbYellow.Text = "80% <= Rate < 90%";
                            lbRed.Text = "70% <= Rate < 80%";
                            lbBlack.Text = "< 70%";
                        }
                    }

                    /////// Get Month
                    if (_dtMonth != null && _dtMonth.Rows.Count > 0)
                    {
                        _last_1_month = _dtMonth.Rows[0]["LAST_1MONTH"].ToString();
                        _last_2_month = _dtMonth.Rows[0]["LAST_2MONTH"].ToString();
                    }

                    /////// Disable Save Button
                    if (_dtCheck != null && _dtCheck.Rows.Count > 0)
                    {
                        _allow_confirm = _dtCheck.Rows[0]["CONFIRM_YN"].ToString().Equals("Y") ? false : true;
                        btnConfirm.Enabled = _dtCheck.Rows[0]["CONFIRM_YN"].ToString().Equals("Y") ? false : true;
                    }

                    if (_dtSource != null && _dtSource.Rows.Count > 0)
                    {
                        for (int iRow = 0; iRow < _dtSource.Rows.Count; iRow++)
                        {
                            if (!string.IsNullOrEmpty(_dtSource.Rows[iRow]["GRP_NAME_VN"].ToString()))
                            {
                                byte[] data = Convert.FromBase64String(_dtSource.Rows[iRow]["GRP_NAME_VN"].ToString());
                                string _txt_viet = Encoding.UTF8.GetString(data);
                                _dtSource.Rows[iRow]["GRP_NAME_VN"] = _txt_viet;
                            }

                            if (!string.IsNullOrEmpty(_dtSource.Rows[iRow]["ITEM_NAME_VN"].ToString()))
                            {
                                byte[] data = Convert.FromBase64String(_dtSource.Rows[iRow]["ITEM_NAME_VN"].ToString());
                                string _txt_viet = Encoding.UTF8.GetString(data);
                                _dtSource.Rows[iRow]["ITEM_NAME_VN"] = _txt_viet;
                            }

                            if (!string.IsNullOrEmpty(_dtSource.Rows[iRow]["CATEGORY_NAME_VN"].ToString()))
                            {
                                byte[] data = Convert.FromBase64String(_dtSource.Rows[iRow]["CATEGORY_NAME_VN"].ToString());
                                string _txt_viet = Encoding.UTF8.GetString(data);
                                _dtSource.Rows[iRow]["CATEGORY_NAME_VN"] = _txt_viet;
                            }

                            if (!string.IsNullOrEmpty(_dtSource.Rows[iRow]["METHOD_NAME_VN"].ToString()))
                            {
                                byte[] data = Convert.FromBase64String(_dtSource.Rows[iRow]["METHOD_NAME_VN"].ToString());
                                string _txt_viet = Encoding.UTF8.GetString(data);
                                _dtSource.Rows[iRow]["METHOD_NAME_VN"] = _txt_viet;
                            }

                            if (!string.IsNullOrEmpty(_dtSource.Rows[iRow]["COUNTERMEASURE"].ToString()))
                            {
                                byte[] data = Convert.FromBase64String(_dtSource.Rows[iRow]["COUNTERMEASURE"].ToString());
                                string _txt_viet = Encoding.UTF8.GetString(data);
                                string _txt_new = _txt_viet.Replace("@", "\n");
                                _dtSource.Rows[iRow]["COUNTERMEASURE"] = _txt_new;
                            }

                            if (_dtSource.Rows[iRow]["ITEM_CD"].ToString().ToUpper().Equals("TOTAL"))
                            {
                                _dtSource.Rows[iRow]["ITEM_NAME_VN"] = "Total";

                                if (_is_other && iRow != _dtSource.Rows.Count - 1)
                                {
                                    _dtSource.Rows[iRow]["CATEGORY_NAME_VN"] = "Total";
                                }
                            }

                            if (_dtSource.Rows[iRow]["GRP_CD"].ToString().ToUpper().Equals("G-TOTAL"))
                            {
                                _dtSource.Rows[iRow]["ITEM_NAME_VN"] = "";
                                _dtSource.Rows[iRow]["GRP_NAME_VN"] = "G-Total";
                            }

                            if(!_dtSource.Rows[iRow]["ITEM_CD"].ToString().ToUpper().Equals("TOTAL") &&
                                _dtSource.Rows[iRow]["RESULT_SCORE"].ToString().Equals("0") &&
                                _dtSource.Rows[iRow]["LINK_YN"].ToString().Equals("N"))
                            {
                                _dtSource.Rows[iRow]["RESULT_SCORE"] = DBNull.Value;
                            }

                            if(iRow.Equals(_dtSource.Rows.Count - 1))
                            {
                                _dtSource.Rows[iRow]["COUNTERMEASURE"] = FormatNumber(_dtSource.Rows[iRow]["RESULT_RATE"].ToString()) + "%";
                            }

                            //if (!string.IsNullOrEmpty(_dtSource.Rows[iRow]["RESULT_TYPE"].ToString()))
                            //{
                            //    string _result_type = _dtSource.Rows[iRow]["RESULT_TYPE"].ToString();

                            //    switch (_result_type)
                            //    {
                            //        case "QTY":
                            //            _dtSource.Rows[iRow]["COUNTERMEASURE"] = "- Result: " + FormatNumber(_dtSource.Rows[iRow]["RESULT"].ToString()) + FormatText(_dtSource.Rows[iRow]["RESULT_UNIT"].ToString());
                            //            break;
                            //        case "RATE":
                            //            _dtSource.Rows[iRow]["COUNTERMEASURE"] = "- Target: " + FormatNumber(_dtSource.Rows[iRow]["TARGET"].ToString()) + FormatText(_dtSource.Rows[iRow]["RESULT_UNIT"].ToString())
                            //                                                   + "\n- Result: " + FormatNumber(_dtSource.Rows[iRow]["QTY"].ToString()) + FormatText(_dtSource.Rows[iRow]["RESULT_UNIT"].ToString())
                            //                                                   + "\n- Rate: " + FormatNumber(_dtSource.Rows[iRow]["RATE"].ToString()) + "%";
                            //            break;
                            //        case "MONTH_RATE":
                            //            _dtSource.Rows[iRow]["COUNTERMEASURE"] = "- Tháng " + _last_2_month + ": " + FormatNumber(_dtSource.Rows[iRow]["TARGET"].ToString()) + FormatText(_dtSource.Rows[iRow]["RESULT_UNIT"].ToString())
                            //                                                   + "\n- Tháng " + _last_1_month + ": " + FormatNumber(_dtSource.Rows[iRow]["QTY"].ToString()) + FormatText(_dtSource.Rows[iRow]["RESULT_UNIT"].ToString())
                            //                                                   + "\n- Rate: " + FormatNumber(_dtSource.Rows[iRow]["RATE"].ToString()) + "%";
                            //            break;
                            //        default:
                            //            break;
                            //    }
                            //}
                        }

                        CreateDetailGrid(grdDetail, gvwDetail, _dtSource);
                        SetData(grdDetail, _dtSource);
                        Formart_Grid_Detail();
                    }
                    else
                    {
                        grdDetail.DataSource = null;
                        gvwDetail.Columns.Clear();
                        gvwDetail.Bands.Clear();
                    }
                }
            }
            catch(Exception ex) { MessageBox.Show(ex.Message); }
            finally
            {
                pbProgressHide();
            }
        }

        public DataTable Binding_Data(DataTable dtSource)
        {
            try
            {
                DataTable _dtf = GetDataTable(gvwSummary);
                string _col_nm = "", _distinct_row = "";

                for (int iRow = 0; iRow < dtSource.Rows.Count; iRow++)
                {
                    if (!dtSource.Rows[iRow]["DIV"].ToString().Equals(_distinct_row))
                    {
                        _dtf.Rows.Add();
                        _dtf.Rows[_dtf.Rows.Count - 1]["DIV"] = dtSource.Rows[iRow]["DIV"].ToString();

                        _distinct_row = dtSource.Rows[iRow]["DIV"].ToString();
                    }

                    _col_nm = dtSource.Rows[iRow]["LINE_CD"].ToString();
                    _dtf.Rows[_dtf.Rows.Count - 1][_col_nm] = dtSource.Rows[iRow]["QTY"].ToString();
                }

                return _dtf;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        public void fn_load_chart(DataTable _dtSource)
        {
            try
            {
                chartData.DataSource = _dtSource;
                chartData.AnnotationRepository.Clear();

                while (chartData.Series[0].Points.Count > 0)
                {
                    chartData.Series[0].Points.Clear();
                }

                XYDiagram diagOSD = (XYDiagram)chartData.Diagram;
                diagOSD.AxisY.WholeRange.MaxValue = 100;
                AxisX axisX = diagOSD.AxisX;
                axisX.CustomLabels.Clear();

                ConstantLine constantLine1 = diagOSD.AxisY.ConstantLines[0];
                constantLine1.AxisValueSerializable = _dtSource.Rows[0]["TARGET"].ToString();

                for (int i = 0; i < _dtSource.Rows.Count; i++)
                {
                    string label = _dtSource.Rows[i]["LINE_NM"].ToString();
                    chartData.Series[0].Points.Add(new SeriesPoint(label, _dtSource.Rows[i]["QTY"].ToString()));
                }

                for (int iRow = 0; iRow < _dtSource.Rows.Count; iRow++)
                {
                    double _rateQty = Double.Parse(_dtSource.Rows[iRow]["QTY"].ToString().Replace("%", ""));

                    if (_rateQty >= 90)
                    {
                        chartData.Series[0].Points[iRow].Color = Color.Green;
                    }
                    else if (_rateQty >= 80 && _rateQty < 90)
                    {
                        chartData.Series[0].Points[iRow].Color = Color.Gold;
                    }
                    else if (_rateQty >= 70 && _rateQty < 80)
                    {
                        chartData.Series[0].Points[iRow].Color = Color.Red;
                    }
                    else
                    {
                        chartData.Series[0].Points[iRow].Color = Color.Black;
                    }
                }

                //if (!cboPlant.EditValue.ToString().Equals("ALL"))
                //{
                //    for (int i = 0; i < _dtSource.Rows.Count; i++)
                //    {
                //        axisX.CustomLabels.Add(new CustomAxisLabel(name: _dtSource.Rows[i]["LINE_NM"].ToString().Replace("_", "").Trim(), value: _dtSource.Rows[i]["LINE_NM"].ToString())
                //        {
                //            TextColor = Color.FromArgb(255, 50, 50, 50),
                //        });
                //    }
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public string FormatNumber(string value)
        {
            return string.IsNullOrEmpty(value) ? "0" : Math.Round(Double.Parse(value), 1).ToString();
        }

        public string FormatText(string value)
        {
            return string.IsNullOrEmpty(value) ? "" : value;
        }

        public override void SaveClick()
        {
            try
            {
                DialogResult dlr;
                int _result_qty = 0, _max_qty = 0; ;

                DataTable _dtf = BindingData(grdDetail, true, false);
                if (_dtf != null && _dtf.Rows.Count > 0)
                {
                    for (int iRow = 0; iRow < _dtf.Rows.Count; iRow++)
                    {
                        _max_qty = string.IsNullOrEmpty(_dtf.Rows[iRow]["MAX_SCORE"].ToString()) ? 0 : Int32.Parse(_dtf.Rows[iRow]["MAX_SCORE"].ToString());
                        _result_qty = string.IsNullOrEmpty(_dtf.Rows[iRow]["RESULT_SCORE"].ToString()) ? 0 : Int32.Parse(_dtf.Rows[iRow]["RESULT_SCORE"].ToString());

                        if (_result_qty < 0 || _result_qty > _max_qty)
                        {
                            MessageBox.Show("Số điểm thực tế phải trong khoảng từ 0 ~ Điểm tối đa!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }

                dlr = MessageBox.Show("Bạn có muốn Save không?", "Save", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dlr == DialogResult.Yes)
                {
                    bool result = SaveData("Q_SAVE");
                    if (result)
                    {
                        MessageBoxW("Save successfully!", IconType.Information);
                        QueryClick();
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

        public void Formart_Grid_Summary()
        {
            try
            {
                grdSummary.BeginUpdate();

                for (int i = 0; i < gvwSummary.Columns.Count; i++)
                {
                    gvwSummary.Columns[i].OptionsColumn.AllowEdit = false;
                    gvwSummary.Columns[i].OptionsColumn.ReadOnly = true;
                    gvwSummary.Columns[i].OptionsColumn.AllowSort = DefaultBoolean.False;
                    gvwSummary.Columns[i].OptionsColumn.AllowMerge = DefaultBoolean.False;

                    gvwSummary.Columns[i].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gvwSummary.Columns[i].AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gvwSummary.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gvwSummary.Columns[i].AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gvwSummary.Columns[i].AppearanceCell.TextOptions.WordWrap = WordWrap.Wrap;
                    gvwSummary.Columns[i].AppearanceCell.Font = new System.Drawing.Font("Calibri", 12, FontStyle.Regular);

                    if (gvwSummary.Columns[i].FieldName.ToString().Equals("DIV"))
                    {
                        gvwSummary.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
                    }

                    if (!gvwSummary.Columns[i].FieldName.ToString().Equals("DIV") && cboGroup.EditValue.ToString().Equals("G002"))
                    {
                        gvwSummary.Columns[i].Width = 85;
                    }
                }

                gvwSummary.RowHeight = 35;
                grdSummary.EndUpdate();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Formart_Grid_Detail()
        {
            try
            {
                grdDetail.BeginUpdate();

                for (int i = 0; i < gvwDetail.Columns.Count; i++)
                {
                    gvwDetail.Columns[i].OptionsColumn.AllowEdit = false;
                    gvwDetail.Columns[i].OptionsColumn.ReadOnly = true;
                    gvwDetail.Columns[i].OptionsColumn.AllowSort = DefaultBoolean.False;

                    gvwDetail.Columns[i].AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gvwDetail.Columns[i].AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gvwDetail.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                    gvwDetail.Columns[i].AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                    gvwDetail.Columns[i].AppearanceCell.TextOptions.WordWrap = WordWrap.Wrap;
                    gvwDetail.Columns[i].AppearanceCell.Font = new System.Drawing.Font("Calibri", 12, FontStyle.Regular);

                    if (gvwDetail.Columns[i].FieldName.ToString().Equals("GRP_NAME_VN"))
                    {
                        gvwDetail.Columns[i].Width = 130;
                        gvwDetail.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                        gvwDetail.Columns[i].ColumnEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
                    }

                    if (gvwDetail.Columns[i].FieldName.ToString().Equals("ORD_NO"))
                    {
                        gvwDetail.Columns[i].Width = 70;
                    }

                    if (gvwDetail.Columns[i].FieldName.ToString().Equals("ITEM_NAME_VN"))
                    {
                        gvwDetail.Columns[i].Width = 100;
                        gvwDetail.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                        gvwDetail.Columns[i].ColumnEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
                    }

                    if (gvwDetail.Columns[i].FieldName.ToString().Equals("CATEGORY_NAME_VN"))
                    {
                        gvwDetail.Columns[i].Width = 280;
                        gvwDetail.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
                        gvwDetail.Columns[i].ColumnEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
                    }

                    if (gvwDetail.Columns[i].FieldName.ToString().Contains("METHOD_NAME_VN"))
                    {
                        gvwDetail.Columns[i].Width = 360;
                        gvwDetail.Columns[i].AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
                        gvwDetail.Columns[i].ColumnEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();
                    }

                    if (gvwDetail.Columns[i].FieldName.ToString().Contains("COUNTERMEASURE"))
                    {
                        gvwDetail.Columns[i].Width = 280;
                        gvwDetail.Columns[i].ColumnEdit = new DevExpress.XtraEditors.Repository.RepositoryItemMemoEdit();

                        if (_allow_confirm)
                        {
                            gvwDetail.Columns[i].OptionsColumn.AllowEdit = true;
                            gvwDetail.Columns[i].OptionsColumn.ReadOnly = false;
                        }
                    }

                    if (gvwDetail.Columns[i].FieldName.ToString().Contains("RESULT_SCORE"))
                    {
                        gvwDetail.Columns[i].Width = 110;

                        if (_allow_confirm)
                        {
                            gvwDetail.Columns[i].OptionsColumn.AllowEdit = true;
                            gvwDetail.Columns[i].OptionsColumn.ReadOnly = false;
                        }
                    }

                    if (gvwDetail.Columns[i].FieldName.ToString().Contains("MAX_SCORE"))
                    {
                        gvwDetail.Columns[i].Width = 100;
                    }

                    if (gvwDetail.Columns[i].FieldName.ToString().Equals("PHOTO"))
                    {
                        gvwDetail.Columns[i].Width = 150;
                        DevExpress.XtraEditors.Repository.RepositoryItemPictureEdit _reposPicture = new DevExpress.XtraEditors.Repository.RepositoryItemPictureEdit();
                        _reposPicture.SizeMode = DevExpress.XtraEditors.Controls.PictureSizeMode.Stretch;
                        gvwDetail.Columns[i].ColumnEdit = _reposPicture;
                    }

                    if (gvwDetail.Columns[i].FieldName.ToString().Equals("AFTER_PHOTO"))
                    {
                        gvwDetail.Columns[i].Width = 200;
                        DevExpress.XtraEditors.Repository.RepositoryItemPictureEdit _reposPicture = new DevExpress.XtraEditors.Repository.RepositoryItemPictureEdit();
                        _reposPicture.SizeMode = DevExpress.XtraEditors.Controls.PictureSizeMode.Stretch;
                        gvwDetail.Columns[i].ColumnEdit = _reposPicture;
                    }
                }

                gvwDetail.RowHeight = 120;
                grdDetail.EndUpdate();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        #endregion [Start Button Event Code By UIBuilder] 

        #region [Grid]

        public void CreateDetailGrid(GridControlEx gridControl, BandedGridViewEx gridView, DataTable dtSource)
        {
            //gridControl.Hide();
            gridView.BeginDataUpdate();
            try
            {
                if (_tab.Equals(0))
                {
                    gridControl.DataSource = null;
                    InitControls(gridControl);
                    gridView.Columns.Clear();
                    gridView.Bands.Clear();

                    while (gridView.Columns.Count > 0)
                    {
                        gridView.Columns.RemoveAt(0);
                    }
                    gridView.OptionsView.ShowColumnHeaders = false;

                    GridBandEx gridBand = null;
                    BandedGridColumnEx colBand = new BandedGridColumnEx();

                    gridBand = new GridBandEx() { Caption = dtSource.Rows[0]["GRP_NM"].ToString() };
                    gridView.Bands.Add(gridBand);
                    gridBand.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                    gridBand.AppearanceHeader.TextOptions.WordWrap = WordWrap.Wrap;
                    gridBand.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center;
                    gridBand.AppearanceHeader.Options.UseBackColor = true;

                    colBand = new BandedGridColumnEx() { FieldName = "DIV", Visible = true };
                    colBand.Width = 100;
                    gridBand.Columns.Add(colBand);

                    for (int iRow = 0; iRow < dtSource.Rows.Count; iRow++)
                    {
                        gridBand = new GridBandEx() { Caption = dtSource.Rows[iRow]["LINE_NM"].ToString() };
                        gridView.Bands.Add(gridBand);
                        gridBand.AppearanceHeader.TextOptions.WordWrap = WordWrap.Wrap;
                        gridBand.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center;
                        gridBand.AppearanceHeader.Options.UseBackColor = true;
                        gridBand.RowCount = 2;

                        colBand = new BandedGridColumnEx() { FieldName = dtSource.Rows[iRow]["LINE_CD"].ToString(), Visible = true };
                        colBand.Width = 75;
                        gridBand.Columns.Add(colBand);
                    }
                }
                else if (_tab.Equals(1))
                {
                    gridControl.DataSource = null;
                    InitControls(gridControl);
                    gridView.Columns.Clear();
                    gridView.Bands.Clear();

                    while (gridView.Columns.Count > 0)
                    {
                        gridView.Columns.RemoveAt(0);
                    }
                    gridView.OptionsView.ShowColumnHeaders = false;

                    GridBandEx gridBand = null;
                    BandedGridColumnEx colBand = new BandedGridColumnEx();
                    int _col_start = Int32.Parse(dtSource.Rows[0]["FIXED_COL_START"].ToString());
                    int _col_end = Int32.Parse(dtSource.Rows[0]["FIXED_COL_END"].ToString());

                    string[] _col_caption = {"Category Group\n(Nhóm hạng mục)","Category\n(Danh mục)", "Category Details\n(Chi tiết danh mục)", "Test Method\n(Phương pháp kiểm tra)",
                        "No\n(STT)", "Max Score\n(Điểm tối đa)", "Actual Score\n(Điểm thực tế)", "Photo\n(Hình ảnh vấn đề)", "Remark & Countermeasure\n(Ghi chú vấn đề & Biện pháp khắc phục)", "After Photo\n(Hình ảnh khắc phục vấn đề)"};
                    string[] _col_field = { "GRP_NAME_VN", "ITEM_NAME_VN", "CATEGORY_NAME_VN", "METHOD_NAME_VN", "ORD_NO", "MAX_SCORE", "RESULT_SCORE", "PHOTO", "COUNTERMEASURE", "AFTER_PHOTO" };

                    for (int iRow = 0; iRow < dtSource.Columns.Count; iRow++)
                    {
                        ////////// KPI Column
                        int iDx = Array.IndexOf(_col_field, dtSource.Columns[iRow].ColumnName.ToString());
                        gridBand = new GridBandEx() { Caption = iDx >= 0 ? _col_caption[iDx] : dtSource.Columns[iRow].ColumnName.ToString() };
                        gridView.Bands.Add(gridBand);

                        if (iRow >= _col_start && iRow <= _col_end)
                        {
                            gridBand.Fixed = DevExpress.XtraGrid.Columns.FixedStyle.Left;
                        }

                        gridBand.AppearanceHeader.TextOptions.WordWrap = WordWrap.Wrap;
                        gridBand.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center;
                        gridBand.AppearanceHeader.Options.UseBackColor = true;
                        gridBand.RowCount = 2;

                        if (_is_other && dtSource.Columns[iRow].ColumnName.ToString().Equals("ITEM_NAME_VN"))
                        {
                            gridBand.Visible = false;
                        }
                        else
                        {
                            gridBand.Visible = _col_field.Contains(dtSource.Columns[iRow].ColumnName.ToString()) ? true : false;
                        }

                        colBand = new BandedGridColumnEx() { FieldName = dtSource.Columns[iRow].ColumnName.ToString(), Visible = _col_field.Contains(dtSource.Columns[iRow].ColumnName.ToString()) };
                        colBand.Width = 120;
                        gridBand.Columns.Add(colBand);
                    }
                }
            }
            catch
            {
                //throw EX;
            }
            //gridControl.Show();
            gridView.EndDataUpdate();
            gridView.ExpandAllGroups();
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
                else if (argType.Equals("Q_SUMMARY") || argType.Equals("Q_SUMMARY_CHART") || argType.Equals("Q_EXPORT"))
                {
                    string _group = cboGroup.EditValue == null ? "" : cboGroup.EditValue.ToString();
                    dtData = proc.SetParamData(dtData, argType, _group, "", "", cboMonth.yyyymm);
                }
                else
                {
                    string _factory = cboFactory.EditValue == null ? "" : cboFactory.EditValue.ToString();
                    string _plant = cboPlant.EditValue == null ? "" : cboPlant.EditValue.ToString();
                    string _line = _null_mline ? "000" : cboLine.EditValue == null ? "" : cboLine.EditValue.ToString();

                    dtData = proc.SetParamData(dtData, argType, _factory, _plant, _line, cboDate.yyyymmdd);
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

        private void toolTipController1_GetActiveObjectInfo(object sender, ToolTipControllerGetActiveObjectInfoEventArgs e)
        {
            try
            {
                if (e.SelectedControl != grdDetail)
                    return;

                GridHitInfo hitInfo = gvwDetail.CalcHitInfo(e.ControlMousePosition);
                if (hitInfo.InRow == false)
                    return;
                if (hitInfo.Column == null) return;

                if (hitInfo.Column.FieldName.Contains("PHOTO") || hitInfo.Column.FieldName.Contains("AFTER_PHOTO"))
                {
                    string _column_nm = hitInfo.Column.FieldName.ToString();

                    if (gvwDetail.GetRowCellValue(hitInfo.RowHandle, _column_nm).ToString() != "")
                    {
                        SuperToolTipSetupArgs toolTipArgs = new SuperToolTipSetupArgs();
                        toolTipArgs.Title.Text = _column_nm == "PHOTO" ? "BEFORE PHOTO" : "AFTER PHOTO";
                        toolTipArgs.Title.Appearance.ForeColor = Color.Silver;
                        e.Info = new ToolTipControlInfo();
                        e.Info.Object = hitInfo.HitTest.ToString() + hitInfo.RowHandle.ToString();
                        e.Info.ToolTipType = ToolTipType.SuperTip;
                        e.Info.SuperTip = new SuperToolTip();
                        e.Info.SuperTip.Padding = new System.Windows.Forms.Padding(1);
                        e.Info.ImmediateToolTip = true;

                        using (WebClient client = new WebClient())
                        {
                            byte[] imageData = (byte[])gvwDetail.GetRowCellValue(hitInfo.RowHandle, _column_nm); // Assuming "image_blob" is the column name

                            using (var ms = new MemoryStream(imageData))
                            {
                                Image image = Image.FromStream(ms);
                                foreach (var prop in image.PropertyItems)
                                {
                                    if ((prop.Id == 0x0112 || prop.Id == 5029 || prop.Id == 274))
                                    {
                                        var value = (int)prop.Value[0];
                                        if (value == 6)
                                        {
                                            image.RotateFlip(RotateFlipType.Rotate90FlipNone);
                                            break;
                                        }
                                        else if (value == 8)
                                        {
                                            image.RotateFlip(RotateFlipType.Rotate270FlipNone);
                                            break;
                                        }
                                        else if (value == 3)
                                        {
                                            image.RotateFlip(RotateFlipType.Rotate180FlipNone);
                                            break;
                                        }
                                    }
                                }

                                toolTipArgs.Contents.Image = ScaleThumbnailImage(image, 400, 300);
                                e.Info.ToolTipImage = image;
                            }
                        }
                        e.Info.SuperTip.Setup(toolTipArgs);
                    }
                }
                else if (hitInfo.Column.FieldName.Contains("CATEGORY_NAME_VN") && _format_cd.Equals("F001"))
                {
                    if(gvwDetail.GetRowCellValue(hitInfo.RowHandle, "GRP_CD").ToString().Equals("G002") && 
                        gvwDetail.GetRowCellValue(hitInfo.RowHandle, "ITEM_CD").ToString().Equals("I006"))
                    {
                        SuperToolTipSetupArgs toolTipArgs = new SuperToolTipSetupArgs();
                        toolTipArgs.Title.Text = "Standard Inventory";
                        toolTipArgs.Title.Appearance.ForeColor = Color.Silver;
                        e.Info = new ToolTipControlInfo();
                        e.Info.Object = hitInfo.HitTest.ToString() + hitInfo.RowHandle.ToString();
                        e.Info.ToolTipType = ToolTipType.SuperTip;
                        e.Info.SuperTip = new SuperToolTip();
                        e.Info.SuperTip.Padding = new System.Windows.Forms.Padding(1);
                        e.Info.ImmediateToolTip = true;

                        using (WebClient client = new WebClient())
                        {
                            byte[] imageData = Convert.FromBase64String(base64String);

                            using (var ms = new MemoryStream(imageData))
                            {
                                Image image = Image.FromStream(ms);
                                foreach (var prop in image.PropertyItems)
                                {
                                    if ((prop.Id == 0x0112 || prop.Id == 5029 || prop.Id == 274))
                                    {
                                        var value = (int)prop.Value[0];
                                        if (value == 6)
                                        {
                                            image.RotateFlip(RotateFlipType.Rotate90FlipNone);
                                            break;
                                        }
                                        else if (value == 8)
                                        {
                                            image.RotateFlip(RotateFlipType.Rotate270FlipNone);
                                            break;
                                        }
                                        else if (value == 3)
                                        {
                                            image.RotateFlip(RotateFlipType.Rotate180FlipNone);
                                            break;
                                        }
                                    }
                                }

                                toolTipArgs.Contents.Image = ScaleThumbnailImage(image, 400, 300);
                                e.Info.ToolTipImage = image;
                            }
                        }
                        e.Info.SuperTip.Setup(toolTipArgs);
                    }
                }
            }
            catch
            {

            }
        }

        internal Image ScaleThumbnailImage(Image ImageToScale, int MaxWidth, int MaxHeight)
        {
            double ratioX = (double)MaxWidth / ImageToScale.Width;
            double ratioY = (double)MaxHeight / ImageToScale.Height;
            double ratio = Math.Min(ratioX, ratioY);

            int newWidth = (int)(ImageToScale.Width * ratio);
            int newHeight = (int)(ImageToScale.Height * ratio);

            Image newImage = new Bitmap(newWidth, newHeight);
            Graphics.FromImage(newImage).DrawImage(ImageToScale, 0, 0, newWidth, newHeight);

            return newImage;
        }

        #endregion [Grid]

        #region [Combobox]

        private void InitCombobox()
        {
            LoadDataCbo(cboFactory, "Factory", "Q_FTY");
            LoadDataCbo(cboPlant, "Plant", "Q_LINE");
            LoadDataCbo(cboLine, "Line", "Q_MLINE");
            LoadDataCbo(cboGroup, "Group", "Q_GROUP");
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

                        lbGreen.Location = new Point(220, 40);
                        lbYellow.Location = new Point(382, 40);
                        lbRed.Location = new Point(529, 40);
                        lbBlack.Location = new Point(676, 40);
                    }

                    return;
                }

                if (_type.Equals("Q_MLINE"))
                {
                    lbLine.Visible = true;
                    cboLine.Visible = true;
                    _null_mline = false;

                    lbGreen.Location = new Point(420, 40);
                    lbYellow.Location = new Point(582, 40);
                    lbRed.Location = new Point(729, 40);
                    lbBlack.Location = new Point(876, 40);
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

        #endregion [Combobox]

        #region Events

        private void tabControl_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            try
            {
                int _prev_tab = _tab;
                _tab = tabControl.SelectedTabPageIndex;
                FormatLayout();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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

        private void gvwDetail_RowCellClick(object sender, RowCellClickEventArgs e)
        {
            try
            {
                if (grdDetail.DataSource == null || gvwDetail.RowCount < 1) return;

                if (!gvwDetail.GetRowCellValue(e.RowHandle, "ITEM_CD").ToString().ToUpper().Equals("TOTAL") &&
                    (e.Column.FieldName.ToString().Equals("PHOTO") || e.Column.FieldName.ToString().Equals("AFTER_PHOTO")))
                {
                    if (e.Clicks >= 2)
                    {
                        if (_allow_confirm)
                        {
                            openFileDialog.FileName = "";
                            openFileDialog.Filter = "Image Files(*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;.bmp;";

                            if (openFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                string _stype = e.Column.FieldName.ToString().Equals("PHOTO") ? "Q_PHOTO" : "Q_AFTER_PHOTO";
                                bool result = SaveData(_stype);
                                if (result)
                                {
                                    MessageBoxW("Upload Photo successfully!", IconType.Information);
                                    QueryClick();
                                }
                                else
                                {
                                    MessageBoxW("Upload failed!", IconType.Warning);
                                }
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private void gvwDetail_CellMerge(object sender, CellMergeEventArgs e)
        {
            try
            {
                if (grdDetail.DataSource == null || gvwDetail.RowCount <= 0) return;

                e.Merge = false;
                e.Handled = true;

                if (e.Column.FieldName.ToString().Equals("GRP_NAME_VN"))
                {
                    string _value1 = gvwDetail.GetRowCellValue(e.RowHandle1, e.Column.FieldName.ToString()).ToString();
                    string _value2 = gvwDetail.GetRowCellValue(e.RowHandle2, e.Column.FieldName.ToString()).ToString();

                    if (_value1 == _value2 && !string.IsNullOrEmpty(_value1))
                    {
                        e.Merge = true;
                    }
                }
            }
            catch
            {

            }
        }

        private void gvwDetail_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            try
            {
                if (grdDetail.DataSource == null || gvwDetail.RowCount < 1) return;

                if (e.Column.FieldName.ToString().Contains("RESULT_SCORE"))
                {
                    e.Appearance.BackColor = Color.FromArgb(226, 239, 217);
                }

                if (gvwDetail.GetRowCellValue(e.RowHandle, "ITEM_CD").ToString().ToUpper().Equals("TOTAL"))
                {
                    e.Appearance.BackColor = Color.LightYellow;
                }

                if (gvwDetail.GetRowCellValue(e.RowHandle, "GRP_CD").ToString().ToUpper().Equals("G-TOTAL"))
                {
                    e.Appearance.BackColor = Color.FromArgb(255, 228, 225);
                    e.Appearance.ForeColor = Color.Blue;
                }

                if ((e.Column.FieldName.ToString().Contains("METHOD_NAME_VN") || e.Column.FieldName.ToString().Equals("CATEGORY_NAME_VN")) && gvwDetail.GetRowCellValue(e.RowHandle, "LINK_YN").ToString().Equals("Y"))
                {
                    e.Appearance.BackColor = Color.FromArgb(237, 237, 237);
                }

                if (gvwDetail.GetRowCellValue(e.RowHandle, "GRP_CD").ToString().ToUpper().Equals("G-TOTAL"))
                {
                    if (e.Column.FieldName.ToString().Equals("COUNTERMEASURE") && !string.IsNullOrEmpty(e.CellValue.ToString()))
                    {
                        double _rateQty = Double.Parse(e.CellValue.ToString().Replace("%",""));
                        if (_is_other)
                        {
                            if (_rateQty >= 90)
                            {
                                e.Appearance.BackColor = Color.Green;
                                e.Appearance.ForeColor = Color.White;
                            }
                            else if (_rateQty >= 85 && _rateQty < 90)
                            {
                                e.Appearance.BackColor = Color.Yellow;
                            }
                            else if (_rateQty >= 80 && _rateQty < 85)
                            {
                                e.Appearance.BackColor = Color.Red;
                                e.Appearance.ForeColor = Color.White;
                            }
                            else
                            {
                                e.Appearance.BackColor = Color.Black;
                                e.Appearance.ForeColor = Color.White;
                            }
                        }
                        else
                        {
                            if (_rateQty >= 90)
                            {
                                e.Appearance.BackColor = Color.Green;
                                e.Appearance.ForeColor = Color.White;
                            }
                            else if (_rateQty >= 80 && _rateQty < 90)
                            {
                                e.Appearance.BackColor = Color.Yellow;
                            }
                            else if (_rateQty >= 70 && _rateQty < 80)
                            {
                                e.Appearance.BackColor = Color.Red;
                                e.Appearance.ForeColor = Color.White;
                            }
                            else
                            {
                                e.Appearance.BackColor = Color.Black;
                                e.Appearance.ForeColor = Color.White;
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private void gvwDetail_ShowingEditor(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                if (grdDetail.DataSource == null || gvwDetail.RowCount < 1) return;

                if (gvwDetail.FocusedRowHandle == gvwDetail.RowCount - 1)
                {
                    e.Cancel = true;
                }

                if (gvwDetail.GetRowCellValue(gvwDetail.FocusedRowHandle, "ITEM_CD").ToString().ToUpper().Equals("TOTAL"))
                {
                    e.Cancel = true;
                }
            }
            catch { }
        }

        private void gvwDetail_CalcRowHeight(object sender, RowHeightEventArgs e)
        {
            try
            {
                if (grdDetail.DataSource == null || gvwDetail.RowCount < 1) return;

                if (e.RowHandle == gvwDetail.RowCount - 1 || gvwDetail.GetRowCellValue(e.RowHandle, "ITEM_CD").ToString().ToUpper().Equals("TOTAL"))
                {
                    e.RowHeight = 45;
                }
            }
            catch { }
        }

        public bool SaveData(string _type)
        {
            try
            {
                bool _result = true;
                DataTable dtData = null;
                P_MSPD90229A_S proc = new P_MSPD90229A_S();
                string machineName = $"{SessionInfo.UserName}|{ Environment.MachineName}|{GetIPAddress()}";
                int iUpdate = 0, iCount = 0;

                pbProgressShow();

                switch (_type)
                {
                    case "Q_SAVE":
                        DataTable _dtf = BindingData(grdDetail, true, false);

                        for (int iRow = 0; iRow < _dtf.Rows.Count; iRow++)
                        {
                            iUpdate++;

                            byte[] dataCounter = System.Text.Encoding.UTF8.GetBytes(_dtf.Rows[iRow]["COUNTERMEASURE"].ToString().Trim());
                            string _txt_counter = System.Convert.ToBase64String(dataCounter);

                            dtData = proc.SetParamData(dtData,
                                                 _type,
                                                 _dtf.Rows[iRow]["PLANT_CD"].ToString(),
                                                 _dtf.Rows[iRow]["LINE_CD"].ToString(),
                                                 _dtf.Rows[iRow]["MLINE_CD"].ToString(),
                                                 cboDate.yyyymmdd,
                                                 _dtf.Rows[iRow]["GRP_CD"].ToString(),
                                                 _dtf.Rows[iRow]["ITEM_CD"].ToString(),
                                                 _dtf.Rows[iRow]["MAX_SCORE"].ToString(),
                                                 _dtf.Rows[iRow]["RESULT_SCORE"].ToString(),
                                                 null,
                                                 _txt_counter,
                                                 machineName,
                                                 "CSI.GMES.PD.MSPD90229A_S");

                            if (CommonProcessSave(dtData, proc.ProcName, proc.GetParamInfo(), grdDetail))
                            {
                                dtData = null;
                                iCount++;
                            }
                            else
                            {
                                break;
                            }
                        }

                        //if (!_is_other)
                        //{
                        //    ///// Save Row Auto Link Data
                        //    dtData = null;
                        //    DataTable _dtSource = GetDataTable(gvwDetail);
                        //    DataTable _dtLink = _dtSource.Select("LINK_YN = 'Y'", "").CopyToDataTable();

                        //    for (int iRow = 0; iRow < _dtLink.Rows.Count; iRow++)
                        //    {
                        //        iUpdate++;

                        //        byte[] dataCounter = System.Text.Encoding.UTF8.GetBytes(_dtLink.Rows[iRow]["COUNTERMEASURE"].ToString().Trim());
                        //        string _txt_counter = System.Convert.ToBase64String(dataCounter);

                        //        dtData = proc.SetParamData(dtData,
                        //                             _type,
                        //                             _dtLink.Rows[iRow]["PLANT_CD"].ToString(),
                        //                             _dtLink.Rows[iRow]["LINE_CD"].ToString(),
                        //                             _dtLink.Rows[iRow]["MLINE_CD"].ToString(),
                        //                             cboDate.yyyymmdd,
                        //                             _dtLink.Rows[iRow]["GRP_CD"].ToString(),
                        //                             _dtLink.Rows[iRow]["ITEM_CD"].ToString(),
                        //                             _dtLink.Rows[iRow]["MAX_SCORE"].ToString(),
                        //                             _dtLink.Rows[iRow]["RESULT_SCORE"].ToString(),
                        //                             null,
                        //                             _txt_counter,
                        //                             machineName,
                        //                             "CSI.GMES.PD.MSPD90229A_S");

                        //        if (CommonProcessSave(dtData, proc.ProcName, proc.GetParamInfo(), grdDetail))
                        //        {
                        //            dtData = null;
                        //            iCount++;
                        //        }
                        //        else
                        //        {
                        //            break;
                        //        }
                        //    }
                        //}

                        if (iUpdate == iCount)
                        {
                            _result = true;
                        }
                        else
                        {
                            _result = false;
                        }

                        break;
                    case "Q_PHOTO":
                    case "Q_AFTER_PHOTO":
                        dtData = proc.SetParamData(dtData,
                                                  _type,
                                                  gvwDetail.GetRowCellValue(gvwDetail.FocusedRowHandle, "PLANT_CD").ToString(),
                                                  gvwDetail.GetRowCellValue(gvwDetail.FocusedRowHandle, "LINE_CD").ToString(),
                                                  gvwDetail.GetRowCellValue(gvwDetail.FocusedRowHandle, "MLINE_CD").ToString(),
                                                  cboDate.yyyymmdd,
                                                  gvwDetail.GetRowCellValue(gvwDetail.FocusedRowHandle, "GRP_CD").ToString(),
                                                  gvwDetail.GetRowCellValue(gvwDetail.FocusedRowHandle, "ITEM_CD").ToString(),
                                                  gvwDetail.GetRowCellValue(gvwDetail.FocusedRowHandle, "MAX_SCORE").ToString(),
                                                  "",
                                                  File.ReadAllBytes(openFileDialog.FileName),
                                                  "",
                                                  machineName,
                                                  "CSI.GMES.PD.MSPD90229A_Q");
                        _result = CommonProcessSave(dtData, proc.ProcName, proc.GetParamInfo(), null);
                        break;
                    case "Q_CONFIRM":
                        DataTable _dtSource = GetDataTable(gvwDetail);
                        DataTable _dtLink = _dtSource.Select("METHOD_NAME_VN IS NOT NULL", "").CopyToDataTable();

                        for (int iRow = 0; iRow < _dtLink.Rows.Count; iRow++)
                        {
                            iUpdate++;

                            byte[] dataCounter = System.Text.Encoding.UTF8.GetBytes(_dtLink.Rows[iRow]["COUNTERMEASURE"].ToString().Trim());
                            string _txt_counter = System.Convert.ToBase64String(dataCounter);

                            dtData = proc.SetParamData(dtData,
                                                 _type,
                                                 _dtLink.Rows[iRow]["PLANT_CD"].ToString(),
                                                 _dtLink.Rows[iRow]["LINE_CD"].ToString(),
                                                 _dtLink.Rows[iRow]["MLINE_CD"].ToString(),
                                                 cboDate.yyyymmdd,
                                                 _dtLink.Rows[iRow]["GRP_CD"].ToString(),
                                                 _dtLink.Rows[iRow]["ITEM_CD"].ToString(),
                                                 _dtLink.Rows[iRow]["MAX_SCORE"].ToString(),
                                                 _dtLink.Rows[iRow]["RESULT_SCORE"].ToString(),
                                                 null,
                                                 _txt_counter,
                                                 machineName,
                                                 "CSI.GMES.PD.MSPD90229A_S");

                            if (CommonProcessSave(dtData, proc.ProcName, proc.GetParamInfo(), grdDetail))
                            {
                                dtData = null;
                                iCount++;
                            }
                            else
                            {
                                break;
                            }
                        }

                        if (iUpdate == iCount)
                        {
                            _result = true;
                        }
                        else
                        {
                            _result = false;
                        }
                        break;
                    default:
                        break;
                }

                pbProgressHide();
                return _result;
            }
            catch (Exception ex)
            {
                pbProgressHide();
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            try
            {
                if (grdDetail.DataSource == null || gvwDetail.RowCount < 1) return;

                DialogResult dlr;
                int _result_qty = 0, _max_qty = 0; ;

                DataTable _dtPermiss = GetData("Q_PERMISS");
                if(_dtPermiss == null || _dtPermiss.Rows.Count < 1)
                {
                    MessageBox.Show("Bạn không có quyền thực hiện chức năng này!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DataTable _dtf = GetDataTable(gvwDetail);
                DataTable _dtLink = _dtf.Select("METHOD_NAME_VN IS NOT NULL", "").CopyToDataTable();

                if (_dtLink != null && _dtLink.Rows.Count > 0)
                {
                    for (int iRow = 0; iRow < _dtLink.Rows.Count; iRow++)
                    {
                        _max_qty = string.IsNullOrEmpty(_dtLink.Rows[iRow]["MAX_SCORE"].ToString()) ? 0 : Int32.Parse(_dtLink.Rows[iRow]["MAX_SCORE"].ToString());
                        _result_qty = string.IsNullOrEmpty(_dtLink.Rows[iRow]["RESULT_SCORE"].ToString()) ? 0 : Int32.Parse(_dtLink.Rows[iRow]["RESULT_SCORE"].ToString());

                        if (_result_qty < 0 || _result_qty > _max_qty)
                        {
                            MessageBox.Show("Số điểm thực tế phải trong khoảng từ 0 ~ Điểm tối đa!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }

                dlr = MessageBox.Show("Bạn có muốn Confirm không?\nLưu ý: Dữ liệu sau khi xác nhận sẽ không được cập nhập nữa!!!", "Save", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dlr == DialogResult.Yes)
                {
                    bool result = SaveData("Q_CONFIRM");
                    if (result)
                    {
                        MessageBoxW("Save successfully!", IconType.Information);
                        QueryClick();
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

        private void gvwSummary_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            try
            {
                if (grdSummary.DataSource == null || gvwSummary.RowCount < 1) return;

                if (!e.Column.FieldName.ToString().Contains("DIV") && e.RowHandle == gvwSummary.RowCount - 2)
                {
                    if (!string.IsNullOrEmpty(e.CellValue.ToString()))
                    {
                        e.DisplayText = FormatNumber(e.CellValue.ToString()) + "%";
                    }
                }
            }
            catch { }
        }

        private void gvwSummary_RowCellStyle(object sender, RowCellStyleEventArgs e)
        {
            try
            {
                if (grdSummary.DataSource == null || gvwSummary.RowCount < 1) return;

                if (e.RowHandle == gvwSummary.RowCount - 1 && !e.Column.FieldName.ToString().Contains("DIV"))
                {
                    e.Appearance.BackColor = Color.LightYellow;

                    if (!string.IsNullOrEmpty(e.CellValue.ToString())){
                        double _rateQty = Double.Parse(e.CellValue.ToString());

                        if(_rateQty == 1)
                        {
                            e.Appearance.BackColor = Color.Green;
                            e.Appearance.ForeColor = Color.White;
                        }
                    }
                }

                if(e.RowHandle == gvwSummary.RowCount - 2)
                {
                    if(!e.Column.FieldName.ToString().Equals("DIV") && !string.IsNullOrEmpty(e.CellValue.ToString()))
                    {
                        double _rateQty = Double.Parse(e.CellValue.ToString().Replace("%", ""));

                        if (_rateQty >= 90)
                        {
                            e.Appearance.BackColor = Color.Green;
                            e.Appearance.ForeColor = Color.White;
                        }
                        else if (_rateQty >= 80 && _rateQty < 90)
                        {
                            e.Appearance.BackColor = Color.Yellow;
                        }
                        else if (_rateQty >= 70 && _rateQty < 80)
                        {
                            e.Appearance.BackColor = Color.Red;
                            e.Appearance.ForeColor = Color.White;
                        }
                        else
                        {
                            e.Appearance.BackColor = Color.Black;
                            e.Appearance.ForeColor = Color.White;
                        }
                    }
                }
            }
            catch { }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable _dtPermiss = GetData("Q_PERMISS");
                if (_dtPermiss == null || _dtPermiss.Rows.Count < 1)
                {
                    MessageBox.Show("Bạn không có quyền thực hiện chức năng này!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DialogResult dlr = MessageBox.Show("Bạn có muốn Send Email không?", "Save", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dlr == DialogResult.Yes)
                {
                    pbProgressShow();

                    P_MSPD90229A_EMAIL proc = new P_MSPD90229A_EMAIL();
                    DataTable dtDataSet = null;
                    dtDataSet = proc.SetParamData(dtDataSet, "Q", cboMonth.yyyymm);

                    ResultSet rs = CommonCallQuery(dtDataSet, proc.ProcName, proc.GetParamInfo(), false, 90000, "", true);
                    if (rs == null || rs.ResultDataSet == null || rs.ResultDataSet.Tables.Count == 0 || rs.ResultDataSet.Tables[0].Rows.Count == 0)
                    {
                        MessageBoxW("Send failed!", IconType.Warning);
                        return;
                    }

                    DataSet dsData = rs.ResultDataSet;
                    DataTable dtData = dsData.Tables[0];
                    DataTable dtChart = dsData.Tables[1];
                    DataTable dtAvg = dsData.Tables[2];
                    DataTable dtHtml = dsData.Tables[3];

                    if (dtData.Rows.Count == 0)
                    {
                        MessageBoxW("Send failed!", IconType.Warning);
                        return;
                    }

                    string SubjectName = "LEAN Foundation Final Result";
                    string htmlReturn = GetHtml(dtData, dtHtml, dtAvg);
                    if (htmlReturn == "") return;

                    /////Image List
                    List<string> imgList = new List<string>();
                    string[] _col_field = { "G001", "G002", "G003" };

                    for (int iCount = 0; iCount < _col_field.Length; iCount++)
                    {
                        DataTable dtGroup = dtChart.Select("OPTION_CD = '" + _col_field[iCount] + "'", "").CopyToDataTable();
                        string picName = "";

                        if (dtGroup != null && dtGroup.Rows.Count > 0)
                        {
                            bool bChart1 = LoadDataChart(dtGroup);
                            if (!bChart1) return;
                            picName = "LEAN_" + _col_field[iCount];
                            CaptureControl(tlpMain, picName);
                            imgList.Add(picName);
                        }
                    }

                    bool _result = CreateMail(SubjectName, htmlReturn, imgList, "", "", "huynh.it@changshininc.com");
                    pbProgressHide();

                    if (_result)
                    {
                        MessageBoxW("Send successfully!", IconType.Information);
                    }
                    else
                    {
                        MessageBoxW("Send failed!", IconType.Warning);
                    }
                }
            }
            catch
            {
                pbProgressHide();
                throw;
            }
        }

        public bool LoadDataChart(DataTable argDt)
        {
            try
            {
                DataTable _dtSource = argDt.Copy();

                chart1.DataSource = argDt;
                chart1.AnnotationRepository.Clear();

                while (chart1.Series[0].Points.Count > 0)
                {
                    chart1.Series[0].Points.Clear();
                }

                XYDiagram diagOSD = (XYDiagram)chart1.Diagram;
                diagOSD.AxisY.WholeRange.MaxValue = 95;
                AxisX axisX = diagOSD.AxisX;
                axisX.CustomLabels.Clear();

                ConstantLine constantLine1 = diagOSD.AxisY.ConstantLines[0];
                constantLine1.AxisValueSerializable = argDt.Rows[0]["TARGET"].ToString();

                for (int i = 0; i < argDt.Rows.Count; i++)
                {
                    string label = argDt.Rows[i]["LINE_NM"].ToString();
                    chart1.Series[0].Points.Add(new SeriesPoint(label, argDt.Rows[i]["QTY"].ToString()));
                }

                for (int iRow = 0; iRow < argDt.Rows.Count; iRow++)
                {
                    string _rateQty = argDt.Rows[iRow]["STATUS"].ToString();

                    switch (_rateQty)
                    {
                        case "GREEN":
                            chart1.Series[0].Points[iRow].Color = Color.Green;
                            break;
                        case "YELLOW":
                            chart1.Series[0].Points[iRow].Color = Color.Gold;
                            break;
                        case "RED":
                            chart1.Series[0].Points[iRow].Color = Color.Red;
                            break;
                        case "BLACK":
                            chart1.Series[0].Points[iRow].Color = Color.Black;
                            break;
                        default:
                            break;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public void CaptureControl(Control control, string nameImg)
        {
            try
            {
                string Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\GMES_LFA_Capture\";
                Bitmap bmp = new Bitmap(control.Width, control.Height);
                if (!Directory.Exists(Path)) Directory.CreateDirectory(Path);
                control.DrawToBitmap(bmp, new Rectangle(0, 0, control.Width, control.Height));
                bmp.Save(Path + nameImg + @".png", System.Drawing.Imaging.ImageFormat.Png);
            }
            catch
            {

            }
        }

        private bool CreateMail(string Subject, string htmlBody, List<string> imgList, string RecipEmail, string MailCC, string MailBCC)
        {
            try
            {
                bool _result = true;

                Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.MailItem mailItem = (Microsoft.Office.Interop.Outlook.MailItem)app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                mailItem.Subject = Subject;
                Microsoft.Office.Interop.Outlook.Recipients oRecips = (Microsoft.Office.Interop.Outlook.Recipients)mailItem.Recipients;

                if (!string.IsNullOrEmpty(RecipEmail))
                {
                    for (int i = 0; i < RecipEmail.Split(';').Length; i++)
                    {
                        Microsoft.Office.Interop.Outlook.Recipient oRecip = (Microsoft.Office.Interop.Outlook.Recipient)oRecips.Add(RecipEmail.Split(';')[i]);
                        oRecip.Resolve();
                    }
                }

                if (MailCC.Trim() != "")
                {
                    mailItem.CC = MailCC;
                }
                if (MailBCC.Trim() != "")
                {
                    mailItem.BCC = MailBCC;
                }

                ////Add Picture
                if (imgList != null)
                {
                    int iPicCount = imgList.Count;
                    string[] imgInfo = new string[iPicCount];
                    StringBuilder strImg = new StringBuilder();
                    string pathPic = "";
                    for (int i = 0; i < iPicCount; i++)
                    {
                        strImg = new StringBuilder();
                        imgInfo[i] = "imgInfo" + (i + 1).ToString();
                        string pic = imgList[i];

                        if (pic.Contains("\\"))
                        {
                            pathPic = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + pic;
                        }
                        else
                        {
                            pic = pic.Contains(".") ? pic : pic + ".png";
                            pathPic = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + $@"\GMES_LFA_Capture\{pic}";
                        }

                        Outlook.Attachment oAttachPic = mailItem.Attachments.Add(pathPic, Outlook.OlAttachmentType.olByValue, null, "tr");
                        oAttachPic.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", imgInfo[i]);
                        strImg.Append("<br><img src='cid:" + imgInfo[i] + "'>");

                        htmlBody = htmlBody.Replace("{chart" + (i + 1) + "}", strImg.ToString());
                    }
                    mailItem.HTMLBody = htmlBody;
                }
                else
                {
                    mailItem.HTMLBody = htmlBody;
                }

                mailItem.HTMLBody = htmlBody;
                mailItem.Importance = Microsoft.Office.Interop.Outlook.OlImportance.olImportanceHigh;
                mailItem.Send();

                return _result;
            }
            catch 
            {
                return false;
            }
        }

        private string GetHtml(DataTable arg_DtData, DataTable arg_DtHtml, DataTable arg_Avg)
        {
            try
            {
                string htmlReturn = arg_DtHtml.Rows[0]["TEXT1"].ToString();
                htmlReturn = htmlReturn.Replace("{MONTH}", arg_DtData.Rows[0]["YYYYMM"].ToString());
                htmlReturn = htmlReturn.Replace("{MMYYYY}", arg_DtData.Rows[0]["MMYYYY"].ToString());

                string htmlTotal = "";

                for (int iRow = 0; iRow < arg_Avg.Rows.Count; iRow++)
                {
                    string htmlAvg = arg_DtHtml.Rows[0]["TEXT7"].ToString();
                    htmlAvg = htmlAvg.Replace("{GRP_NM}", arg_Avg.Rows[iRow]["GRP_NM"].ToString());
                    htmlAvg = htmlAvg.Replace("{LINE_LIST}", arg_Avg.Rows[iRow]["LINE_LIST"].ToString());
                    htmlAvg = htmlAvg.Replace("{AVG_GROUP}", arg_Avg.Rows[iRow]["TOTAL_AVG"].ToString());
                    htmlAvg = htmlAvg.Replace("{AVG_LINE}", arg_Avg.Rows[iRow]["QTY"].ToString());

                    htmlTotal += htmlAvg;
                }
                htmlReturn = htmlReturn.Replace("{ttotal}", htmlTotal);

                string _distinct_row = "", _columnHtml = "", _distinct_div = "", _row_html = "";
                string _headType = "", _bodyType = "";

                for (int iRow = 0; iRow < arg_DtData.Rows.Count; iRow++)
                {
                    if (!_distinct_row.Equals(arg_DtData.Rows[iRow]["OPTION_CD"].ToString()))
                    {
                        _distinct_row = arg_DtData.Rows[iRow]["OPTION_CD"].ToString();

                        string htmlTableHead = arg_DtHtml.Rows[0]["TEXT2"].ToString();
                        htmlTableHead = htmlTableHead.Replace("{GRP_NM}", arg_DtData.Rows[iRow]["GRP_NM"].ToString());

                        DataTable dtGroup = arg_DtData.Select("OPTION_CD = '" + _distinct_row + "'", "").CopyToDataTable();

                        if (dtGroup != null && dtGroup.Rows.Count > 0)
                        {
                            /////Table Header
                            var distinctValues = dtGroup.AsEnumerable()
                                .Select(row => new
                                {
                                    ORD = row.Field<decimal>("ORD"),
                                    LINE_CD = row.Field<string>("LINE_CD"),
                                    LINE_NM = row.Field<string>("LINE_NM"),
                                })
                                 .Distinct().OrderBy(r => r.ORD);
                            DataTable _dtHead = LINQResultToDataTable(distinctValues).Select("", "").CopyToDataTable();
                            string htmlHead = "";

                            for (int iCount = 0; iCount < _dtHead.Rows.Count; iCount++)
                            {
                                _columnHtml = arg_DtHtml.Rows[0]["TEXT3"].ToString();
                                _columnHtml = _columnHtml.Replace("{LINE_NM}", string.Format("{0}", _dtHead.Rows[iCount]["LINE_NM"].ToString()));
                                htmlHead += _columnHtml;
                            }

                            htmlTableHead = htmlTableHead.Replace("{tPlant}", htmlHead);

                            ///////Table Row
                            string htmlOther = "", htmlCell = "";
                            string htmlTableRow = "";

                            for (int iCount = 0; iCount < dtGroup.Rows.Count; iCount++)
                            {
                                if (!_distinct_div.Equals(dtGroup.Rows[iCount]["DIV"].ToString()))
                                {
                                    _distinct_div = dtGroup.Rows[iCount]["DIV"].ToString();

                                    if (!string.IsNullOrEmpty(htmlOther) && !string.IsNullOrEmpty(_row_html))
                                    {
                                        _row_html = _row_html.Replace("{tRow}", htmlOther);
                                        htmlTableRow += _row_html;
                                    }

                                    htmlOther = "";
                                    _row_html = arg_DtHtml.Rows[0]["TEXT4"].ToString();
                                    _row_html = _row_html.Replace("{DIV}", dtGroup.Rows[iCount]["DIV"].ToString());
                                }

                                if (_distinct_div.ToUpper().Equals("RANK"))
                                {
                                    htmlCell = arg_DtHtml.Rows[0]["TEXT6"].ToString();
                                    htmlCell = htmlCell.Replace("{ALIGN}", "center");
                                }
                                else
                                {
                                    htmlCell = arg_DtHtml.Rows[0]["TEXT5"].ToString();
                                    if (_distinct_div.Contains("%"))
                                    {
                                        htmlCell = htmlCell.Replace("{ALIGN}", "center");
                                    }
                                    else
                                    {
                                        htmlCell = htmlCell.Replace("{ALIGN}", "right");
                                    }
                                }

                                htmlCell = htmlCell.Replace("{QTY}", dtGroup.Rows[iCount]["QTY"].ToString());
                                htmlCell = htmlCell.Replace("{COLOR}", dtGroup.Rows[iCount]["STATUS"].ToString());

                                htmlOther += htmlCell;

                                if (iCount.Equals(dtGroup.Rows.Count - 1))
                                {
                                    _row_html = _row_html.Replace("{tRow}", htmlOther);
                                    htmlTableRow += _row_html;
                                }
                            }

                            ////
                            switch (_distinct_row)
                            {
                                case "G001":
                                    _headType = "thead1";
                                    _bodyType = "tbody1";
                                    break;
                                case "G002":
                                    _headType = "thead2";
                                    _bodyType = "tbody2";
                                    break;
                                case "G003":
                                    _headType = "thead3";
                                    _bodyType = "tbody3";
                                    break;
                                default:
                                    break;
                            }
                            htmlReturn = htmlReturn.Replace("{" + _headType + "}", htmlTableHead);
                            htmlReturn = htmlReturn.Replace("{" + _bodyType + "}", htmlTableRow);
                        }
                    }
                }

                return htmlReturn;
            }
            catch (Exception ex)
            {
                return "";
            }
        }

        public void FormatNumber_Excel(ExcelRange range)
        {
            foreach (var cell in range)
            {
                string stringValue = cell.Value?.ToString(); // Assuming the cell contains a string value
                double numericValue;
                if (double.TryParse(stringValue, out numericValue))
                {
                    cell.Value = numericValue; // Assign the numeric value to the cell
                }
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

        public class P_MSPD90229A_EMAIL : BaseProcClass
        {
            public P_MSPD90229A_EMAIL()
            {
                // Modify Code : Procedure Name
                _ProcName = "LMES.P_MSPD90229A_EMAIL";
                ParamAdd();
            }
            private void ParamAdd()
            {
                _ParamInfo.Add(new ParamInfo("@V_P_SEND_TYPE", "Varchar", 100, "Input", typeof(System.String)));
                _ParamInfo.Add(new ParamInfo("@V_P_SEND_DATE", "Varchar", 100, "Input", typeof(System.String)));
            }
            public DataTable SetParamData(DataTable dataTable,
                                        System.String V_P_SEND_TYPE,
                                        System.String V_P_SEND_DATE)
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
                                                V_P_SEND_TYPE,
                                                V_P_SEND_DATE
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

        DataTable GetDataTable(GridView view)
        {
            DataTable dt = new DataTable();
            foreach (GridColumn c in view.Columns)
                dt.Columns.Add(c.FieldName, c.ColumnType);
            for (int r = 0; r < view.RowCount; r++)
            {
                object[] rowValues = new object[dt.Columns.Count];
                for (int c = 0; c < dt.Columns.Count; c++)
                    rowValues[c] = view.GetRowCellValue(r, dt.Columns[c].ColumnName);
                dt.Rows.Add(rowValues);
            }
            return dt;
        }

        private DataTable LINQResultToDataTable<T>(IEnumerable<T> Linqlist)
        {
            DataTable dt = new DataTable();
            PropertyInfo[] columns = null;
            if (Linqlist == null) return dt;
            foreach (T Record in Linqlist)
            {
                if (columns == null)
                {
                    columns = ((Type)Record.GetType()).GetProperties();
                    foreach (PropertyInfo GetProperty in columns)
                    {
                        Type colType = GetProperty.PropertyType;

                        if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition()
                        == typeof(Nullable<>)))
                        {
                            colType = colType.GetGenericArguments()[0];
                        }

                        dt.Columns.Add(new DataColumn(GetProperty.Name, colType));
                    }
                }
                DataRow dr = dt.NewRow();
                foreach (PropertyInfo pinfo in columns)
                {
                    dr[pinfo.Name] = pinfo.GetValue(Record, null) == null ? DBNull.Value : pinfo.GetValue
                    (Record, null);
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}