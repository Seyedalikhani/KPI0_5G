using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
//using System.Linq;

namespace KPI0_5G
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        public Excel.Application xlApp1 { get; set; }
        public Excel.Application xlApp2 { get; set; }
        public Excel.Workbook workbook1 { get; set; }
        public Excel.Workbook workbook2 { get; set; }

        public Excel.Worksheet sheet1 { get; set; }
        public Excel.Worksheet sheet2 { get; set; }
        public Excel.Worksheet Site_sheet { get; set; }

        string[] Site_Vec = new string[1000];
        int Site_Vec_Ind = 0;

        private void button1_Click(object sender, EventArgs e)
        {


            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel File|*.xlsx";
            DialogResult result = openFileDialog1.ShowDialog();
            string File_Name = openFileDialog1.SafeFileName.ToString();
            if (result == DialogResult.OK)
            {

                string file = openFileDialog1.FileName;
                xlApp1 = new Excel.Application();
                workbook1 = xlApp1.Workbooks.Open(file);
                sheet1 = workbook1.Worksheets[1];
                sheet2 = workbook1.Worksheets[2];




                Excel.Range History_TT = sheet1.get_Range("A2", "Q" + sheet1.UsedRange.Rows.Count);
                object[,] FARAZ_Data = (object[,])History_TT.Value;
                int Count = sheet1.UsedRange.Rows.Count;

                Excel.Range History_TT2 = sheet2.get_Range("A2", "D" + sheet2.UsedRange.Rows.Count);
                object[,] FARAZ_Data2 = (object[,])History_TT2.Value;
                int Count2 = sheet2.UsedRange.Rows.Count;

                for (int k = 0; k < Count - 1; k++)
                {

                    string Cell = FARAZ_Data[k + 1, 2].ToString();
                    string Site = Cell.Substring(5, 2) + Cell.Substring(9, 4);

                    if (!Site_Vec.Contains(Site))
                    {
                        Site_Vec[Site_Vec_Ind] = Site;
                        Site_Vec_Ind++;

                        var xlNewSheet = (Excel.Worksheet)workbook1.Worksheets.Add(Type.Missing, workbook1.Worksheets[2], Type.Missing, Type.Missing);
                        xlNewSheet.Name = Site;
                    }


                }


                for (int k = 0; k < Site_Vec_Ind; k++)
                {


                    Site_sheet = workbook1.Worksheets[2 + k + 1];

                    Site_sheet.Cells[1, 1] = "Technology";
                    Site_sheet.Cells[2, 1] = "NR";
                    Site_sheet.Cells[3, 1] = "NR";
                    Site_sheet.Cells[4, 1] = "NR";
                    Site_sheet.Cells[5, 1] = "NR";
                    Site_sheet.Cells[6, 1] = "NR";
                    Site_sheet.Cells[7, 1] = "NR";
                    Site_sheet.Cells[8, 1] = "NR";
                    Site_sheet.Cells[9, 1] = "NR";
                    Site_sheet.Cells[10, 1] = "NR";
                    Site_sheet.Cells[11, 1] = "NR";
                    Site_sheet.Cells[12, 1] = "NR";
                    Site_sheet.Cells[13, 1] = "NR";
                    Site_sheet.Cells[14, 1] = "NR";
                    Site_sheet.Cells[15, 1] = "NR";
                    Site_sheet.Cells[16, 1] = "NR";
                    Site_sheet.Cells[17, 1] = "NR";
                    Site_sheet.Cells[18, 1] = "NR";
                    Site_sheet.Cells[19, 1] = "NR";
                    Site_sheet.Cells[20, 1] = "NR";

                    Site_sheet.Cells[1, 2] = "Category";
                    Site_sheet.Cells[2, 2] = "SCG Add";
                    Site_sheet.Cells[3, 2] = "SCG Change";
                    Site_sheet.Cells[4, 2] = "SCG  Drop Rate";
                    Site_sheet.Cells[5, 2] = "DL traffic GB";
                    Site_sheet.Cells[6, 2] = "UL traffic GB";
                    Site_sheet.Cells[7, 2] = "Rank 2 Usage(%)";
                    Site_sheet.Cells[8, 2] = "Rank 3 Usage(%)";
                    Site_sheet.Cells[9, 2] = "Rank 4 Usage(%)";
                    Site_sheet.Cells[10, 2] = "CQI";
                    Site_sheet.Cells[11, 2] = "MCS";
                    Site_sheet.Cells[12, 2] = "DL User Thp Mbps (FH=23)";
                    Site_sheet.Cells[13, 2] = "DL User Thp Mbps (Daily)";
                    Site_sheet.Cells[14, 2] = "UL User Thp Mbps (FH=23)";
                    Site_sheet.Cells[15, 2] = "UL User Thp Mbps (Daily)";
                    Site_sheet.Cells[16, 2] = "Pcell change Succes Rate";
                    Site_sheet.Cells[17, 2] = "RB_Utilizing Rate DL";
                    Site_sheet.Cells[18, 2] = "RB_Utilizing Rate UL";
                    Site_sheet.Cells[19, 2] = "Cell Unavailable Rate";
                    Site_sheet.Cells[20, 2] = "Average User Number";


                    Site_sheet.Cells[1, 3] = "KPI Name ";
                    Site_sheet.Cells[2, 3] = "EN_DC_Setup_Sucess_Rate_Captured_in_gNodeb%";
                    Site_sheet.Cells[3, 3] = "";
                    Site_sheet.Cells[4, 3] = "Endc_drop_rate%";
                    Site_sheet.Cells[5, 3] = "MAC_DL_Traffic_DRB+SRB_Gbyte";
                    Site_sheet.Cells[6, 3] = "MAC_UL_Traffic_DRB+SRB_Gbyte";
                    Site_sheet.Cells[7, 3] = "rank_2_report%";
                    Site_sheet.Cells[8, 3] = "rank_3_report%";
                    Site_sheet.Cells[9, 3] = "rank_4_report%";
                    Site_sheet.Cells[10, 3] = "Average_CQI_256QAM";
                    Site_sheet.Cells[11, 3] = "Average_MCS_Table2_Downlink_up_to_256QAM";
                    Site_sheet.Cells[12, 3] = "Average_Downlink_MAC_User_Throughput_Mbps";
                    Site_sheet.Cells[13, 3] = "Average_Downlink_MAC_User_Throughput_Mbps";
                    Site_sheet.Cells[14, 3] = "Average_Uplink_MAC_User_Throughput_Mbps";
                    Site_sheet.Cells[15, 3] = "Average_Uplink_MAC_User_Throughput_Mbps";
                    Site_sheet.Cells[16, 3] = "";
                    Site_sheet.Cells[17, 3] = "downlink_resource_block_utilization%";
                    Site_sheet.Cells[18, 3] = "uplink_resource_block_utilization%";
                    Site_sheet.Cells[19, 3] = "100-5G_Cell_Availability_Rate%";
                    Site_sheet.Cells[20, 3] = "Average_Of_Average_Number_of_RRC_Connected_ENDC_NSA";




                    //Site_sheet.Cells[1, 3] = "Faraz Name";
                    //Site_sheet.Cells[2, 3] = "Average of SgNB_Addition_Success_Rate_gNodeb_side%";
                    //Site_sheet.Cells[3, 3] = "Average of SCG_Change_Success_Rate%_Lte_Side";
                    //Site_sheet.Cells[4, 3] = "Average of SgNB_Triggered_SgNB_Abnormal_Release_Rate%";
                    //Site_sheet.Cells[5, 3] = "Sum of Downlink_Traffic_GB_RLC_Layer";
                    //Site_sheet.Cells[6, 3] = "Sum of Uplink_Traffic_GB_RLC_Layer";
                    //Site_sheet.Cells[7, 3] = "Average_MCS_PDSCH_64QAM_Rank2";
                    //Site_sheet.Cells[8, 3] = "Average_MCS_PDSCH_64QAM_Rank3";
                    //Site_sheet.Cells[9, 3] = "Average_MCS_PDSCH_64QAM_Rank4";
                    //Site_sheet.Cells[10, 3] = "Average_CQI_64QAM+256QAM";
                    //Site_sheet.Cells[11, 3] = "Average_MCS_Downlink_pdsch_Huawei_5G_Cell";
                    //Site_sheet.Cells[12, 3] = "Average of Downlink_User_Throughput_Mbps_RLC_Layer";
                    //Site_sheet.Cells[13, 3] = "Average of Downlink_User_Throughput_Mbps_RLC_Layer";
                    //Site_sheet.Cells[14, 3] = "Average of Uplink_User_Throughput_Mbps_RLC_Layer";
                    //Site_sheet.Cells[15, 3] = "Average of Uplink_User_Throughput_Mbps_RLC_Layer";
                    //Site_sheet.Cells[16, 3] = "Average of PCell_Change_Executions_Success_Rate%";
                    //Site_sheet.Cells[17, 3] = "Average of Downlink_Resource_Block_Utilizing_Rate%";
                    //Site_sheet.Cells[18, 3] = "Average of Uplink_Resource_Block_Utilizing_Rate%";
                    //Site_sheet.Cells[19, 3] = "100-(Average of Cell_Availability_Rate_Huawei_5G%/24)";
                    //Site_sheet.Cells[20, 3] = "Average of MAX_Of_Average_Number_of_LTE_NR_NSA_DC_UEs_Huawei_Cell_5G";



                }

                for (int n = 0; n < Site_Vec_Ind; n++)
                {

                    string Site = Site_Vec[n].ToString();
                    Count = sheet1.UsedRange.Rows.Count;
                    int index_of_sheet = Site_Vec_Ind + 2 - n;
                    Count2 = sheet2.UsedRange.Rows.Count;

                    double SCG_Add_A_Value = 0;
                    double SCG_Change_A_Value = 0;
                    double SCG_Drop_Rate_A_Value = 0;
                    double DL_traffic_GB_A_Value = 0;
                    double UL_traffic_GB_A_Value = 0;
                    double Rank_2_Usage_A_Value = 0;
                    double Rank_3_Usage_A_Value = 0;
                    double Rank_4_Usage_A_Value = 0;
                    double CQI_A_Value = 0;
                    double MCS_A_Value = 0;
                    double DL_User_Thp_Mbps_Daily_A_Value = 0;
                    double DL_User_Thp_Mbps_23_A_Value = 0;
                    double UL_User_Thp_Mbps_Daily_A_Value = 0;
                    double UL_User_Thp_Mbps_23_A_Value = 0;
                    double Pcell_change_Succes_Rate_A_Value = 0;
                    double RB_Utilizing_Rate_DL_A_Value = 0;
                    double RB_Utilizing_Rate_UL_A_Value = 0;
                    double Cell_Unavailable_Rate_A_Value = 0;
                    double Average_User_Number_A_Value = 0;




                    int SCG_Add_A = 0;
                    int SCG_Change_A = 0;
                    int SCG_Drop_Rate_A = 0;
                    int DL_traffic_GB_A = 0;
                    int UL_traffic_GB_A = 0;
                    int Rank_2_Usage_A = 0;
                    int Rank_3_Usage_A = 0;
                    int Rank_4_Usage_A = 0;
                    int CQI_A = 0;
                    int MCS_A = 0;
                    int DL_User_Thp_Mbps_Daily_A = 0;
                    int DL_User_Thp_Mbps_23_A = 0;
                    int UL_User_Thp_Mbps_Daily_A = 0;
                    int UL_User_Thp_Mbps_23_A = 0;
                    int Pcell_change_Succes_Rate_A = 0;
                    int RB_Utilizing_Rate_DL_A = 0;
                    int RB_Utilizing_Rate_UL_A = 0;
                    int Cell_Unavailable_Rate_A = 0;
                    int Average_User_Number_A = 0;

                    double SCG_Add_B_Value = 0;
                    double SCG_Change_B_Value = 0;
                    double SCG_Drop_Rate_B_Value = 0;
                    double DL_traffic_GB_B_Value = 0;
                    double UL_traffic_GB_B_Value = 0;
                    double Rank_2_Usage_B_Value = 0;
                    double Rank_3_Usage_B_Value = 0;
                    double Rank_4_Usage_B_Value = 0;
                    double CQI_B_Value = 0;
                    double MCS_B_Value = 0;
                    double DL_User_Thp_Mbps_Daily_B_Value = 0;
                    double DL_User_Thp_Mbps_23_B_Value = 0;
                    double UL_User_Thp_Mbps_Daily_B_Value = 0;
                    double UL_User_Thp_Mbps_23_B_Value = 0;
                    double Pcell_change_Succes_Rate_B_Value = 0;
                    double RB_Utilizing_Rate_DL_B_Value = 0;
                    double RB_Utilizing_Rate_UL_B_Value = 0;
                    double Cell_Unavailable_Rate_B_Value = 0;
                    double Average_User_Number_B_Value = 0;


                    int SCG_Add_B = 0;
                    int SCG_Change_B = 0;
                    int SCG_Drop_Rate_B = 0;
                    int DL_traffic_GB_B = 0;
                    int UL_traffic_GB_B = 0;
                    int Rank_2_Usage_B = 0;
                    int Rank_3_Usage_B = 0;
                    int Rank_4_Usage_B = 0;
                    int CQI_B = 0;
                    int MCS_B = 0;
                    int DL_User_Thp_Mbps_Daily_B = 0;
                    int DL_User_Thp_Mbps_23_B = 0;
                    int UL_User_Thp_Mbps_Daily_B = 0;
                    int UL_User_Thp_Mbps_23_B = 0;
                    int Pcell_change_Succes_Rate_B = 0;
                    int RB_Utilizing_Rate_DL_B = 0;
                    int RB_Utilizing_Rate_UL_B = 0;
                    int Cell_Unavailable_Rate_B = 0;
                    int Average_User_Number_B = 0;



                    double SCG_Add_C_Value = 0;
                    double SCG_Change_C_Value = 0;
                    double SCG_Drop_Rate_C_Value = 0;
                    double DL_traffic_GB_C_Value = 0;
                    double UL_traffic_GB_C_Value = 0;
                    double Rank_2_Usage_C_Value = 0;
                    double Rank_3_Usage_C_Value = 0;
                    double Rank_4_Usage_C_Value = 0;
                    double CQI_C_Value = 0;
                    double MCS_C_Value = 0;
                    double DL_User_Thp_Mbps_Daily_C_Value = 0;
                    double DL_User_Thp_Mbps_23_C_Value = 0;
                    double UL_User_Thp_Mbps_Daily_C_Value = 0;
                    double UL_User_Thp_Mbps_23_C_Value = 0;
                    double Pcell_change_Succes_Rate_C_Value = 0;
                    double RB_Utilizing_Rate_DL_C_Value = 0;
                    double RB_Utilizing_Rate_UL_C_Value = 0;
                    double Cell_Unavailable_Rate_C_Value = 0;
                    double Average_User_Number_C_Value = 0;


                    int SCG_Add_C = 0;
                    int SCG_Change_C = 0;
                    int SCG_Drop_Rate_C = 0;
                    int DL_traffic_GB_C = 0;
                    int UL_traffic_GB_C = 0;
                    int Rank_2_Usage_C = 0;
                    int Rank_3_Usage_C = 0;
                    int Rank_4_Usage_C = 0;
                    int CQI_C = 0;
                    int MCS_C = 0;
                    int DL_User_Thp_Mbps_Daily_C = 0;
                    int DL_User_Thp_Mbps_23_C = 0;
                    int UL_User_Thp_Mbps_Daily_C = 0;
                    int UL_User_Thp_Mbps_23_C = 0;
                    int Pcell_change_Succes_Rate_C = 0;
                    int RB_Utilizing_Rate_DL_C = 0;
                    int RB_Utilizing_Rate_UL_C = 0;
                    int Cell_Unavailable_Rate_C = 0;
                    int Average_User_Number_C = 0;




                    double SCG_Add_D_Value = 0;
                    double SCG_Change_D_Value = 0;
                    double SCG_Drop_Rate_D_Value = 0;
                    double DL_traffic_GB_D_Value = 0;
                    double UL_traffic_GB_D_Value = 0;
                    double Rank_2_Usage_D_Value = 0;
                    double Rank_3_Usage_D_Value = 0;
                    double Rank_4_Usage_D_Value = 0;
                    double CQI_D_Value = 0;
                    double MCS_D_Value = 0;
                    double DL_User_Thp_Mbps_Daily_D_Value = 0;
                    double DL_User_Thp_Mbps_23_D_Value = 0;
                    double UL_User_Thp_Mbps_Daily_D_Value = 0;
                    double UL_User_Thp_Mbps_23_D_Value = 0;
                    double Pcell_change_Succes_Rate_D_Value = 0;
                    double RB_Utilizing_Rate_DL_D_Value = 0;
                    double RB_Utilizing_Rate_UL_D_Value = 0;
                    double Cell_Unavailable_Rate_D_Value = 0;
                    double Average_User_Number_D_Value = 0;


                    int SCG_Add_D = 0;
                    int SCG_Change_D = 0;
                    int SCG_Drop_Rate_D = 0;
                    int DL_traffic_GB_D = 0;
                    int UL_traffic_GB_D = 0;
                    int Rank_2_Usage_D = 0;
                    int Rank_3_Usage_D = 0;
                    int Rank_4_Usage_D = 0;
                    int CQI_D = 0;
                    int MCS_D = 0;
                    int DL_User_Thp_Mbps_Daily_D = 0;
                    int DL_User_Thp_Mbps_23_D = 0;
                    int UL_User_Thp_Mbps_Daily_D = 0;
                    int UL_User_Thp_Mbps_23_D = 0;
                    int Pcell_change_Succes_Rate_D = 0;
                    int RB_Utilizing_Rate_DL_D = 0;
                    int RB_Utilizing_Rate_UL_D = 0;
                    int Cell_Unavailable_Rate_D = 0;
                    int Average_User_Number_D = 0;



                    for (int k = 0; k < Count - 1; k++)
                    {

                        string Cell = FARAZ_Data[k + 1, 2].ToString();
                        string Site1 = Cell.Substring(5, 2) + Cell.Substring(9, 4);

                        if (Site1== Site)
                        {

                            string Sector = Cell.Substring(13, 1);

                            if (Sector=="A")
                            {
                                if (FARAZ_Data[k + 1, 3] != null)
                                {
                                    SCG_Add_A++;
                                    SCG_Add_A_Value = SCG_Add_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 3].ToString());
                                }

                                //string SCG_Change_str = FARAZ_Data[k + 1, 4].ToString();
                                //if (SCG_Change_str != null && SCG_Change_str != "")
                                //{
                                //    SCG_Change_A++;
                                //    SCG_Change = SCG_Change + Convert.ToDouble(FARAZ_Data[k + 1, 4].ToString());
                                //}

                                if (FARAZ_Data[k + 1, 4] != null)
                                {
                                    SCG_Drop_Rate_A++;
                                    SCG_Drop_Rate_A_Value = SCG_Drop_Rate_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 4].ToString());
                                }


                                if (FARAZ_Data[k + 1, 5] != null )
                                {
                                    DL_traffic_GB_A++;
                                    DL_traffic_GB_A_Value = DL_traffic_GB_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 5].ToString());
                                }



                                if (FARAZ_Data[k + 1, 6] != null)
                                {
                                    UL_traffic_GB_A++;
                                    UL_traffic_GB_A_Value = UL_traffic_GB_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 6].ToString());
                                }


                                if (FARAZ_Data[k + 1, 7] != null )
                                {
                                    Rank_2_Usage_A++;
                                    Rank_2_Usage_A_Value = Rank_2_Usage_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 7].ToString());
                                }



                                if (FARAZ_Data[k + 1, 8] != null )
                                {
                                    Rank_3_Usage_A++;
                                    Rank_3_Usage_A_Value = Rank_3_Usage_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 8].ToString());
                                }


                                if (FARAZ_Data[k + 1, 9] != null )
                                {
                                    Rank_4_Usage_A++;
                                    Rank_4_Usage_A_Value = Rank_4_Usage_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 9].ToString());
                                }

                                if (FARAZ_Data[k + 1, 10] != null )
                                {
                                    CQI_A++;
                                    CQI_A_Value = CQI_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 10].ToString());
                                }


                                if (FARAZ_Data[k + 1, 11] != null )
                                {
                                    MCS_A++;
                                    MCS_A_Value = MCS_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 11].ToString());
                                }

                                if (FARAZ_Data[k + 1, 12] != null )
                                {
                                    DL_User_Thp_Mbps_Daily_A++;
                                    DL_User_Thp_Mbps_Daily_A_Value = DL_User_Thp_Mbps_Daily_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 12].ToString());
                                }


                                //string DL_User_Thp_Mbps_23_str = FARAZ_Data[k + 1, 13].ToString();
                                //if (DL_User_Thp_Mbps_23_str != null && DL_User_Thp_Mbps_23_str != "")
                                //{
                                //    DL_User_Thp_Mbps_23_A++;
                                //    DL_User_Thp_Mbps_23 = DL_User_Thp_Mbps_23 + Convert.ToDouble(FARAZ_Data[k + 1, 13].ToString());
                                //}

                                if (FARAZ_Data[k + 1, 13] != null )
                                {
                                    UL_User_Thp_Mbps_Daily_A++;
                                    UL_User_Thp_Mbps_Daily_A_Value = UL_User_Thp_Mbps_Daily_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 13].ToString());
                                }


                                //string UL_User_Thp_Mbps_23_str = FARAZ_Data[k + 1, 15].ToString();
                                //if (UL_User_Thp_Mbps_23_str != null && UL_User_Thp_Mbps_23_str != "")
                                //{
                                //    UL_User_Thp_Mbps_23_A++;
                                //    UL_User_Thp_Mbps_23 = UL_User_Thp_Mbps_23 + Convert.ToDouble(FARAZ_Data[k + 1, 15].ToString());
                                //}

                                //string Pcell_change_Succes_Rate_str = FARAZ_Data[k + 1, 14].ToString();
                                //if (Pcell_change_Succes_Rate_str != null && Pcell_change_Succes_Rate_str != "")
                                //{
                                //    Pcell_change_Succes_Rate_A++;
                                //    Pcell_change_Succes_Rate = Pcell_change_Succes_Rate + Convert.ToDouble(FARAZ_Data[k + 1, 14].ToString());
                                //}

            
                                if (FARAZ_Data[k + 1, 14] != null )
                                {
                                    RB_Utilizing_Rate_DL_A++;
                                    RB_Utilizing_Rate_DL_A_Value = RB_Utilizing_Rate_DL_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 14].ToString());
                                }


                                if (FARAZ_Data[k + 1, 15] != null )
                                {
                                    RB_Utilizing_Rate_UL_A++;
                                    RB_Utilizing_Rate_UL_A_Value = RB_Utilizing_Rate_UL_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 15].ToString());
                                }



                                if (FARAZ_Data[k + 1, 16] != null )
                                {
                                    Cell_Unavailable_Rate_A++;
                                    Cell_Unavailable_Rate_A_Value = Cell_Unavailable_Rate_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 16].ToString());
                                }



                                if (FARAZ_Data[k + 1, 17] != null)
                                {
                                    Average_User_Number_A++;
                                    Average_User_Number_A_Value = Average_User_Number_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 17].ToString());
                                }


                            }





                            if (Sector == "B")
                            {
                                if (FARAZ_Data[k + 1, 3] != null)
                                {
                                    SCG_Add_B++;
                                    SCG_Add_B_Value = SCG_Add_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 3].ToString());
                                }

                                //string SCG_Change_str = FARAZ_Data[k + 1, 4].ToString();
                                //if (SCG_Change_str != null && SCG_Change_str != "")
                                //{
                                //    SCG_Change_A++;
                                //    SCG_Change = SCG_Change + Convert.ToDouble(FARAZ_Data[k + 1, 4].ToString());
                                //}

                                if (FARAZ_Data[k + 1, 4] != null)
                                {
                                    SCG_Drop_Rate_B++;
                                    SCG_Drop_Rate_B_Value = SCG_Drop_Rate_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 4].ToString());
                                }


                                if (FARAZ_Data[k + 1, 5] != null)
                                {
                                    DL_traffic_GB_B++;
                                    DL_traffic_GB_B_Value = DL_traffic_GB_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 5].ToString());
                                }



                                if (FARAZ_Data[k + 1, 6] != null)
                                {
                                    UL_traffic_GB_B++;
                                    UL_traffic_GB_B_Value = UL_traffic_GB_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 6].ToString());
                                }


                                if (FARAZ_Data[k + 1, 7] != null)
                                {
                                    Rank_2_Usage_B++;
                                    Rank_2_Usage_B_Value = Rank_2_Usage_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 7].ToString());
                                }



                                if (FARAZ_Data[k + 1, 8] != null)
                                {
                                    Rank_3_Usage_B++;
                                    Rank_3_Usage_B_Value = Rank_3_Usage_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 8].ToString());
                                }


                                if (FARAZ_Data[k + 1, 9] != null)
                                {
                                    Rank_4_Usage_B++;
                                    Rank_4_Usage_B_Value = Rank_4_Usage_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 9].ToString());
                                }

                                if (FARAZ_Data[k + 1, 10] != null)
                                {
                                    CQI_B++;
                                    CQI_B_Value = CQI_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 10].ToString());
                                }


                                if (FARAZ_Data[k + 1, 11] != null)
                                {
                                    MCS_B++;
                                    MCS_B_Value = MCS_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 11].ToString());
                                }

                                if (FARAZ_Data[k + 1, 12] != null)
                                {
                                    DL_User_Thp_Mbps_Daily_B++;
                                    DL_User_Thp_Mbps_Daily_B_Value = DL_User_Thp_Mbps_Daily_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 12].ToString());
                                }


                                //string DL_User_Thp_Mbps_23_str = FARAZ_Data[k + 1, 13].ToString();
                                //if (DL_User_Thp_Mbps_23_str != null && DL_User_Thp_Mbps_23_str != "")
                                //{
                                //    DL_User_Thp_Mbps_23_A++;
                                //    DL_User_Thp_Mbps_23 = DL_User_Thp_Mbps_23 + Convert.ToDouble(FARAZ_Data[k + 1, 13].ToString());
                                //}

                                if (FARAZ_Data[k + 1, 13] != null)
                                {
                                    UL_User_Thp_Mbps_Daily_B++;
                                    UL_User_Thp_Mbps_Daily_B_Value = UL_User_Thp_Mbps_Daily_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 13].ToString());
                                }


                                //string UL_User_Thp_Mbps_23_str = FARAZ_Data[k + 1, 15].ToString();
                                //if (UL_User_Thp_Mbps_23_str != null && UL_User_Thp_Mbps_23_str != "")
                                //{
                                //    UL_User_Thp_Mbps_23_A++;
                                //    UL_User_Thp_Mbps_23 = UL_User_Thp_Mbps_23 + Convert.ToDouble(FARAZ_Data[k + 1, 15].ToString());
                                //}

                                //string Pcell_change_Succes_Rate_str = FARAZ_Data[k + 1, 14].ToString();
                                //if (Pcell_change_Succes_Rate_str != null && Pcell_change_Succes_Rate_str != "")
                                //{
                                //    Pcell_change_Succes_Rate_A++;
                                //    Pcell_change_Succes_Rate = Pcell_change_Succes_Rate + Convert.ToDouble(FARAZ_Data[k + 1, 14].ToString());
                                //}


                                if (FARAZ_Data[k + 1, 14] != null)
                                {
                                    RB_Utilizing_Rate_DL_B++;
                                    RB_Utilizing_Rate_DL_B_Value = RB_Utilizing_Rate_DL_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 14].ToString());
                                }


                                if (FARAZ_Data[k + 1, 15] != null)
                                {
                                    RB_Utilizing_Rate_UL_B++;
                                    RB_Utilizing_Rate_UL_B_Value = RB_Utilizing_Rate_UL_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 15].ToString());
                                }



                                if (FARAZ_Data[k + 1, 16] != null)
                                {
                                    Cell_Unavailable_Rate_B++;
                                    Cell_Unavailable_Rate_B_Value = Cell_Unavailable_Rate_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 16].ToString());
                                }



                                if (FARAZ_Data[k + 1, 17] != null)
                                {
                                    Average_User_Number_B++;
                                    Average_User_Number_B_Value = Average_User_Number_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 17].ToString());
                                }


                            }






                            if (Sector == "C")
                            {
                                if (FARAZ_Data[k + 1, 3] != null)
                                {
                                    SCG_Add_C++;
                                    SCG_Add_C_Value = SCG_Add_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 3].ToString());
                                }

                                //string SCG_Change_str = FARAZ_Data[k + 1, 4].ToString();
                                //if (SCG_Change_str != null && SCG_Change_str != "")
                                //{
                                //    SCG_Change_A++;
                                //    SCG_Change = SCG_Change + Convert.ToDouble(FARAZ_Data[k + 1, 4].ToString());
                                //}

                                if (FARAZ_Data[k + 1, 4] != null)
                                {
                                    SCG_Drop_Rate_C++;
                                    SCG_Drop_Rate_C_Value = SCG_Drop_Rate_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 4].ToString());
                                }


                                if (FARAZ_Data[k + 1, 5] != null)
                                {
                                    DL_traffic_GB_C++;
                                    DL_traffic_GB_C_Value = DL_traffic_GB_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 5].ToString());
                                }



                                if (FARAZ_Data[k + 1, 6] != null)
                                {
                                    UL_traffic_GB_C++;
                                    UL_traffic_GB_C_Value = UL_traffic_GB_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 6].ToString());
                                }


                                if (FARAZ_Data[k + 1, 7] != null)
                                {
                                    Rank_2_Usage_C++;
                                    Rank_2_Usage_C_Value = Rank_2_Usage_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 7].ToString());
                                }



                                if (FARAZ_Data[k + 1, 8] != null)
                                {
                                    Rank_3_Usage_C++;
                                    Rank_3_Usage_C_Value = Rank_3_Usage_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 8].ToString());
                                }


                                if (FARAZ_Data[k + 1, 9] != null)
                                {
                                    Rank_4_Usage_C++;
                                    Rank_4_Usage_C_Value = Rank_4_Usage_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 9].ToString());
                                }

                                if (FARAZ_Data[k + 1, 10] != null)
                                {
                                    CQI_C++;
                                    CQI_C_Value = CQI_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 10].ToString());
                                }


                                if (FARAZ_Data[k + 1, 11] != null)
                                {
                                    MCS_C++;
                                    MCS_C_Value = MCS_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 11].ToString());
                                }

                                if (FARAZ_Data[k + 1, 12] != null)
                                {
                                    DL_User_Thp_Mbps_Daily_C++;
                                    DL_User_Thp_Mbps_Daily_C_Value = DL_User_Thp_Mbps_Daily_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 12].ToString());
                                }


                                //string DL_User_Thp_Mbps_23_str = FARAZ_Data[k + 1, 13].ToString();
                                //if (DL_User_Thp_Mbps_23_str != null && DL_User_Thp_Mbps_23_str != "")
                                //{
                                //    DL_User_Thp_Mbps_23_A++;
                                //    DL_User_Thp_Mbps_23 = DL_User_Thp_Mbps_23 + Convert.ToDouble(FARAZ_Data[k + 1, 13].ToString());
                                //}

                                if (FARAZ_Data[k + 1, 13] != null)
                                {
                                    UL_User_Thp_Mbps_Daily_C++;
                                    UL_User_Thp_Mbps_Daily_C_Value = UL_User_Thp_Mbps_Daily_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 13].ToString());
                                }


                                //string UL_User_Thp_Mbps_23_str = FARAZ_Data[k + 1, 15].ToString();
                                //if (UL_User_Thp_Mbps_23_str != null && UL_User_Thp_Mbps_23_str != "")
                                //{
                                //    UL_User_Thp_Mbps_23_A++;
                                //    UL_User_Thp_Mbps_23 = UL_User_Thp_Mbps_23 + Convert.ToDouble(FARAZ_Data[k + 1, 15].ToString());
                                //}

                                //string Pcell_change_Succes_Rate_str = FARAZ_Data[k + 1, 14].ToString();
                                //if (Pcell_change_Succes_Rate_str != null && Pcell_change_Succes_Rate_str != "")
                                //{
                                //    Pcell_change_Succes_Rate_A++;
                                //    Pcell_change_Succes_Rate = Pcell_change_Succes_Rate + Convert.ToDouble(FARAZ_Data[k + 1, 14].ToString());
                                //}


                                if (FARAZ_Data[k + 1, 14] != null)
                                {
                                    RB_Utilizing_Rate_DL_C++;
                                    RB_Utilizing_Rate_DL_C_Value = RB_Utilizing_Rate_DL_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 14].ToString());
                                }


                                if (FARAZ_Data[k + 1, 15] != null)
                                {
                                    RB_Utilizing_Rate_UL_C++;
                                    RB_Utilizing_Rate_UL_C_Value = RB_Utilizing_Rate_UL_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 15].ToString());
                                }



                                if (FARAZ_Data[k + 1, 16] != null)
                                {
                                    Cell_Unavailable_Rate_C++;
                                    Cell_Unavailable_Rate_C_Value = Cell_Unavailable_Rate_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 16].ToString());
                                }



                                if (FARAZ_Data[k + 1, 17] != null)
                                {
                                    Average_User_Number_C++;
                                    Average_User_Number_C_Value = Average_User_Number_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 17].ToString());
                                }


                            }




                            if (Sector == "D")
                            {
                                if (FARAZ_Data[k + 1, 3] != null)
                                {
                                    SCG_Add_D++;
                                    SCG_Add_D_Value = SCG_Add_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 3].ToString());
                                }

                                //string SCG_Dhange_str = FARAZ_Data[k + 1, 4].ToString();
                                //if (SCG_Dhange_str != null && SCG_Dhange_str != "")
                                //{
                                //    SCG_Dhange_A++;
                                //    SCG_Dhange = SCG_Dhange + Convert.ToDouble(FARAZ_Data[k + 1, 4].ToString());
                                //}

                                if (FARAZ_Data[k + 1, 4] != null)
                                {
                                    SCG_Drop_Rate_D++;
                                    SCG_Drop_Rate_D_Value = SCG_Drop_Rate_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 4].ToString());
                                }


                                if (FARAZ_Data[k + 1, 5] != null)
                                {
                                    DL_traffic_GB_D++;
                                    DL_traffic_GB_D_Value = DL_traffic_GB_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 5].ToString());
                                }



                                if (FARAZ_Data[k + 1, 6] != null)
                                {
                                    UL_traffic_GB_D++;
                                    UL_traffic_GB_D_Value = UL_traffic_GB_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 6].ToString());
                                }


                                if (FARAZ_Data[k + 1, 7] != null)
                                {
                                    Rank_2_Usage_D++;
                                    Rank_2_Usage_D_Value = Rank_2_Usage_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 7].ToString());
                                }



                                if (FARAZ_Data[k + 1, 8] != null)
                                {
                                    Rank_3_Usage_D++;
                                    Rank_3_Usage_D_Value = Rank_3_Usage_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 8].ToString());
                                }


                                if (FARAZ_Data[k + 1, 9] != null)
                                {
                                    Rank_4_Usage_D++;
                                    Rank_4_Usage_D_Value = Rank_4_Usage_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 9].ToString());
                                }

                                if (FARAZ_Data[k + 1, 10] != null)
                                {
                                    CQI_D++;
                                    CQI_D_Value = CQI_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 10].ToString());
                                }


                                if (FARAZ_Data[k + 1, 11] != null)
                                {
                                    MCS_D++;
                                    MCS_D_Value = MCS_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 11].ToString());
                                }

                                if (FARAZ_Data[k + 1, 12] != null)
                                {
                                    DL_User_Thp_Mbps_Daily_D++;
                                    DL_User_Thp_Mbps_Daily_D_Value = DL_User_Thp_Mbps_Daily_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 12].ToString());
                                }


                                //string DL_User_Thp_Mbps_23_str = FARAZ_Data[k + 1, 13].ToString();
                                //if (DL_User_Thp_Mbps_23_str != null && DL_User_Thp_Mbps_23_str != "")
                                //{
                                //    DL_User_Thp_Mbps_23_A++;
                                //    DL_User_Thp_Mbps_23 = DL_User_Thp_Mbps_23 + Convert.ToDouble(FARAZ_Data[k + 1, 13].ToString());
                                //}

                                if (FARAZ_Data[k + 1, 13] != null)
                                {
                                    UL_User_Thp_Mbps_Daily_D++;
                                    UL_User_Thp_Mbps_Daily_D_Value = UL_User_Thp_Mbps_Daily_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 13].ToString());
                                }


                                //string UL_User_Thp_Mbps_23_str = FARAZ_Data[k + 1, 15].ToString();
                                //if (UL_User_Thp_Mbps_23_str != null && UL_User_Thp_Mbps_23_str != "")
                                //{
                                //    UL_User_Thp_Mbps_23_A++;
                                //    UL_User_Thp_Mbps_23 = UL_User_Thp_Mbps_23 + Convert.ToDouble(FARAZ_Data[k + 1, 15].ToString());
                                //}

                                //string Pcell_Dhange_Succes_Rate_str = FARAZ_Data[k + 1, 14].ToString();
                                //if (Pcell_Dhange_Succes_Rate_str != null && Pcell_Dhange_Succes_Rate_str != "")
                                //{
                                //    Pcell_Dhange_Succes_Rate_A++;
                                //    Pcell_Dhange_Succes_Rate = Pcell_Dhange_Succes_Rate + Convert.ToDouble(FARAZ_Data[k + 1, 14].ToString());
                                //}


                                if (FARAZ_Data[k + 1, 14] != null)
                                {
                                    RB_Utilizing_Rate_DL_D++;
                                    RB_Utilizing_Rate_DL_D_Value = RB_Utilizing_Rate_DL_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 14].ToString());
                                }


                                if (FARAZ_Data[k + 1, 15] != null)
                                {
                                    RB_Utilizing_Rate_UL_D++;
                                    RB_Utilizing_Rate_UL_D_Value = RB_Utilizing_Rate_UL_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 15].ToString());
                                }



                                if (FARAZ_Data[k + 1, 16] != null)
                                {
                                    Cell_Unavailable_Rate_D++;
                                    Cell_Unavailable_Rate_D_Value = Cell_Unavailable_Rate_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 16].ToString());
                                }



                                if (FARAZ_Data[k + 1, 17] != null)
                                {
                                    Average_User_Number_D++;
                                    Average_User_Number_D_Value = Average_User_Number_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 17].ToString());
                                }


                            }



                        }


                    }




                    Site_sheet = workbook1.Worksheets[index_of_sheet];


                    SCG_Add_A_Value = SCG_Add_A_Value / SCG_Add_A;
                    SCG_Change_A_Value = SCG_Change_A_Value / SCG_Change_A;
                    SCG_Drop_Rate_A_Value = SCG_Drop_Rate_A_Value / SCG_Drop_Rate_A;
                    //DL_traffic_GB = DL_traffic_GB ;
                    //UL_traffic_GB = UL_traffic_GB ;
                    Rank_2_Usage_A_Value = Rank_2_Usage_A_Value / Rank_2_Usage_A;
                    Rank_3_Usage_A_Value = Rank_3_Usage_A_Value / Rank_3_Usage_A;
                    Rank_4_Usage_A_Value = Rank_4_Usage_A_Value / Rank_4_Usage_A;
                    CQI_A_Value = CQI_A_Value / CQI_A;
                    MCS_A_Value = MCS_A_Value / MCS_A;
                    DL_User_Thp_Mbps_Daily_A_Value = DL_User_Thp_Mbps_Daily_A_Value / DL_User_Thp_Mbps_Daily_A;
                    //DL_User_Thp_Mbps_23_A_Value = DL_User_Thp_Mbps_23_A_Value / DL_User_Thp_Mbps_23_A;
                    UL_User_Thp_Mbps_Daily_A_Value = UL_User_Thp_Mbps_Daily_A_Value / UL_User_Thp_Mbps_Daily_A;
                    //UL_User_Thp_Mbps_23_A_Value = UL_User_Thp_Mbps_23_A_Value / UL_User_Thp_Mbps_23_A;
                    Pcell_change_Succes_Rate_A_Value = Pcell_change_Succes_Rate_A_Value / Pcell_change_Succes_Rate_A;
                    RB_Utilizing_Rate_DL_A_Value = RB_Utilizing_Rate_DL_A_Value / RB_Utilizing_Rate_DL_A;
                    RB_Utilizing_Rate_UL_A_Value = RB_Utilizing_Rate_UL_A_Value / RB_Utilizing_Rate_UL_A;
                    Cell_Unavailable_Rate_A_Value = Cell_Unavailable_Rate_A_Value / Cell_Unavailable_Rate_A;
                    Average_User_Number_A_Value = Average_User_Number_A_Value / Average_User_Number_A;




                    Site_sheet.Cells[1, 4] = "Sector A";
                    if (SCG_Add_A_Value != 0 && SCG_Add_A != 0)
                    {
                        Site_sheet.Cells[2, 4] = SCG_Add_A_Value;
                    }
                    Site_sheet.Cells[3, 4] = "";
                    if (SCG_Drop_Rate_A_Value != 0 && SCG_Drop_Rate_A != 0)
                    {
                        Site_sheet.Cells[4, 4] = SCG_Drop_Rate_A_Value;
                    }
                    if (DL_traffic_GB_A_Value != 0 && DL_traffic_GB_A != 0)
                    {
                        Site_sheet.Cells[5, 4] = DL_traffic_GB_A_Value;
                    }
                    if (UL_traffic_GB_A_Value != 0 && UL_traffic_GB_A != 0)
                    {
                        Site_sheet.Cells[6, 4] = UL_traffic_GB_A_Value;
                    }
                    if (Rank_2_Usage_A_Value != 0 && Rank_2_Usage_A != 0)
                    {
                        Site_sheet.Cells[7, 4] = Rank_2_Usage_A_Value;
                    }
                    if (Rank_3_Usage_A_Value != 0 && Rank_3_Usage_A != 0)
                    {
                        Site_sheet.Cells[8, 4] = Rank_3_Usage_A_Value;
                    }
                    if (Rank_4_Usage_A_Value != 0 && Rank_4_Usage_A != 0)
                    {
                        Site_sheet.Cells[9, 4] = Rank_4_Usage_A_Value;
                    }
                    if (CQI_A_Value != 0 && CQI_A != 0)
                    {
                        Site_sheet.Cells[10, 4] = CQI_A_Value;
                    }
                    if (MCS_A_Value != 0 && MCS_A != 0)
                    {
                        Site_sheet.Cells[11, 4] = MCS_A_Value;
                    }
                    if (DL_User_Thp_Mbps_Daily_A_Value != 0 && DL_User_Thp_Mbps_Daily_A != 0)
                    {
                        Site_sheet.Cells[13, 4] = DL_User_Thp_Mbps_Daily_A_Value;
                    }
                    //Site_sheet.Cells[13, 4] = DL_User_Thp_Mbps_23_A_Value;
                    if (UL_User_Thp_Mbps_Daily_A_Value != 0 && UL_User_Thp_Mbps_Daily_A != 0)
                    {
                        Site_sheet.Cells[15, 4] = UL_User_Thp_Mbps_Daily_A_Value;
                    }
                    //Site_sheet.Cells[15, 4] = UL_User_Thp_Mbps_23_A_Value;
                    Site_sheet.Cells[16, 4] = "";
                    if (RB_Utilizing_Rate_DL_A_Value != 0 && RB_Utilizing_Rate_DL_A != 0)
                    {
                        Site_sheet.Cells[17, 4] = RB_Utilizing_Rate_DL_A_Value;
                    }
                    if (RB_Utilizing_Rate_UL_A_Value != 0 && RB_Utilizing_Rate_UL_A != 0)
                    {
                        Site_sheet.Cells[18, 4] = RB_Utilizing_Rate_UL_A_Value;
                    }
                    if (Cell_Unavailable_Rate_A_Value != 0 && Cell_Unavailable_Rate_A != 0)
                    {
                        Site_sheet.Cells[19, 4] = 100 - Cell_Unavailable_Rate_A_Value;
                    }
                    if (Average_User_Number_A_Value != 0 && Average_User_Number_A != 0)
                    {
                        Site_sheet.Cells[20, 4] = Average_User_Number_A_Value;
                    }

                    SCG_Add_B_Value = SCG_Add_B_Value / SCG_Add_B;
                    SCG_Change_B_Value = SCG_Change_B_Value / SCG_Change_B;
                    SCG_Drop_Rate_B_Value = SCG_Drop_Rate_B_Value / SCG_Drop_Rate_B;
                    //DL_traffic_GB = DL_traffic_GB ;
                    //UL_traffic_GB = UL_traffic_GB ;
                    Rank_2_Usage_B_Value = Rank_2_Usage_B_Value / Rank_2_Usage_B;
                    Rank_3_Usage_B_Value = Rank_3_Usage_B_Value / Rank_3_Usage_B;
                    Rank_4_Usage_B_Value = Rank_4_Usage_B_Value / Rank_4_Usage_B;
                    CQI_B_Value = CQI_B_Value / CQI_B;
                    MCS_B_Value = MCS_B_Value / MCS_B;
                    DL_User_Thp_Mbps_Daily_B_Value = DL_User_Thp_Mbps_Daily_B_Value / DL_User_Thp_Mbps_Daily_B;
                    //DL_User_Thp_Mbps_23_B_Value = DL_User_Thp_Mbps_23_B_Value / DL_User_Thp_Mbps_23_B;
                    UL_User_Thp_Mbps_Daily_B_Value = UL_User_Thp_Mbps_Daily_B_Value / UL_User_Thp_Mbps_Daily_B;
                    //UL_User_Thp_Mbps_23_B_Value = UL_User_Thp_Mbps_23_B_Value / UL_User_Thp_Mbps_23_B;
                    Pcell_change_Succes_Rate_B_Value = Pcell_change_Succes_Rate_B_Value / Pcell_change_Succes_Rate_B;
                    RB_Utilizing_Rate_DL_B_Value = RB_Utilizing_Rate_DL_B_Value / RB_Utilizing_Rate_DL_B;
                    RB_Utilizing_Rate_UL_B_Value = RB_Utilizing_Rate_UL_B_Value / RB_Utilizing_Rate_UL_B;
                    Cell_Unavailable_Rate_B_Value = Cell_Unavailable_Rate_B_Value / Cell_Unavailable_Rate_B;
                    Average_User_Number_B_Value = Average_User_Number_B_Value / Average_User_Number_B;


                    Site_sheet.Cells[1, 5] = "Sector B";
                    if (SCG_Add_B_Value != 0 && SCG_Add_B != 0)
                    {
                        Site_sheet.Cells[2, 5] = SCG_Add_B_Value;
                    }
                    Site_sheet.Cells[3, 5] = "";
                    if (SCG_Drop_Rate_B_Value != 0 && SCG_Drop_Rate_B != 0)
                    {
                        Site_sheet.Cells[4, 5] = SCG_Drop_Rate_B_Value;
                    }
                    if (DL_traffic_GB_B_Value != 0 && DL_traffic_GB_B != 0)
                    {
                        Site_sheet.Cells[5, 5] = DL_traffic_GB_B_Value;
                    }
                    if (UL_traffic_GB_B_Value != 0 && UL_traffic_GB_B != 0)
                    {
                        Site_sheet.Cells[6, 5] = UL_traffic_GB_B_Value;
                    }
                    if (Rank_2_Usage_B_Value != 0 && Rank_2_Usage_B != 0)
                    {
                        Site_sheet.Cells[7, 5] = Rank_2_Usage_B_Value;
                    }
                    if (Rank_3_Usage_B_Value != 0 && Rank_3_Usage_B != 0)
                    {
                        Site_sheet.Cells[8, 5] = Rank_3_Usage_B_Value;
                    }
                    if (Rank_4_Usage_B_Value != 0 && Rank_4_Usage_B != 0)
                    {
                        Site_sheet.Cells[9, 5] = Rank_4_Usage_B_Value;
                    }
                    if (CQI_B_Value != 0 && CQI_B != 0)
                    {
                        Site_sheet.Cells[10, 5] = CQI_B_Value;
                    }
                    if (MCS_B_Value != 0 && MCS_B != 0)
                    {
                        Site_sheet.Cells[11, 5] = MCS_B_Value;
                    }
                    if (DL_User_Thp_Mbps_Daily_B_Value != 0 && DL_User_Thp_Mbps_Daily_B != 0)
                    {
                        Site_sheet.Cells[13, 5] = DL_User_Thp_Mbps_Daily_B_Value;
                    }
                    //Site_sheet.Cells[13, 5] = DL_User_Thp_Mbps_23_B_Value;
                    if (UL_User_Thp_Mbps_Daily_B_Value != 0 && UL_User_Thp_Mbps_Daily_B != 0)
                    {
                        Site_sheet.Cells[15, 5] = UL_User_Thp_Mbps_Daily_B_Value;
                    }
                    //Site_sheet.Cells[15, 5] = UL_User_Thp_Mbps_23_B_Value;
                    Site_sheet.Cells[16, 5] = "";
                    if (RB_Utilizing_Rate_DL_B_Value != 0 && RB_Utilizing_Rate_DL_B != 0)
                    {
                        Site_sheet.Cells[17, 5] = RB_Utilizing_Rate_DL_B_Value;
                    }
                    if (RB_Utilizing_Rate_UL_B_Value != 0 && RB_Utilizing_Rate_UL_B != 0)
                    {
                        Site_sheet.Cells[18, 5] = RB_Utilizing_Rate_UL_B_Value;
                    }
                    if (Cell_Unavailable_Rate_B_Value != 0 && Cell_Unavailable_Rate_B != 0)
                    {
                        Site_sheet.Cells[19, 5] = 100 - Cell_Unavailable_Rate_B_Value;
                    }
                    if (Average_User_Number_B_Value != 0 && Average_User_Number_B != 0)
                    {
                        Site_sheet.Cells[20, 5] = Average_User_Number_B_Value;
                    }


                    SCG_Add_C_Value = SCG_Add_C_Value / SCG_Add_C;
                    SCG_Change_C_Value = SCG_Change_C_Value / SCG_Change_C;
                    SCG_Drop_Rate_C_Value = SCG_Drop_Rate_C_Value / SCG_Drop_Rate_C;
                    //DL_traffic_GB = DL_traffic_GB ;
                    //UL_traffic_GB = UL_traffic_GB ;
                    Rank_2_Usage_C_Value = Rank_2_Usage_C_Value / Rank_2_Usage_C;
                    Rank_3_Usage_C_Value = Rank_3_Usage_C_Value / Rank_3_Usage_C;
                    Rank_4_Usage_C_Value = Rank_4_Usage_C_Value / Rank_4_Usage_C;
                    CQI_C_Value = CQI_C_Value / CQI_C;
                    MCS_C_Value = MCS_C_Value / MCS_C;
                    DL_User_Thp_Mbps_Daily_C_Value = DL_User_Thp_Mbps_Daily_C_Value / DL_User_Thp_Mbps_Daily_C;
                    //DL_User_Thp_Mbps_23_C_Value = DL_User_Thp_Mbps_23_C_Value / DL_User_Thp_Mbps_23_C;
                    UL_User_Thp_Mbps_Daily_C_Value = UL_User_Thp_Mbps_Daily_C_Value / UL_User_Thp_Mbps_Daily_C;
                    //UL_User_Thp_Mbps_23_C_Value = UL_User_Thp_Mbps_23_C_Value / UL_User_Thp_Mbps_23_C;
                    Pcell_change_Succes_Rate_C_Value = Pcell_change_Succes_Rate_C_Value / Pcell_change_Succes_Rate_C;
                    RB_Utilizing_Rate_DL_C_Value = RB_Utilizing_Rate_DL_C_Value / RB_Utilizing_Rate_DL_C;
                    RB_Utilizing_Rate_UL_C_Value = RB_Utilizing_Rate_UL_C_Value / RB_Utilizing_Rate_UL_C;
                    Cell_Unavailable_Rate_C_Value = Cell_Unavailable_Rate_C_Value / Cell_Unavailable_Rate_C;
                    Average_User_Number_C_Value = Average_User_Number_C_Value / Average_User_Number_C;




                    Site_sheet.Cells[1, 6] = "Sector C";
                    if (SCG_Add_C_Value != 0 && SCG_Add_C != 0)
                    {
                        Site_sheet.Cells[2, 6] = SCG_Add_C_Value;
                    }
                    Site_sheet.Cells[3, 6] = "";
                    if (SCG_Drop_Rate_C_Value != 0 && SCG_Drop_Rate_C != 0)
                    {
                        Site_sheet.Cells[4, 6] = SCG_Drop_Rate_C_Value;
                    }
                    if (DL_traffic_GB_C_Value != 0 && DL_traffic_GB_C != 0)
                    {
                        Site_sheet.Cells[5, 6] = DL_traffic_GB_C_Value;
                    }
                    if (UL_traffic_GB_C_Value != 0 && UL_traffic_GB_C != 0)
                    {
                        Site_sheet.Cells[6, 6] = UL_traffic_GB_C_Value;
                    }
                    if (Rank_2_Usage_C_Value != 0 && Rank_2_Usage_C != 0)
                    {
                        Site_sheet.Cells[7, 6] = Rank_2_Usage_C_Value;
                    }
                    if (Rank_3_Usage_C_Value != 0 && Rank_3_Usage_C != 0)
                    {
                        Site_sheet.Cells[8, 6] = Rank_3_Usage_C_Value;
                    }
                    if (Rank_4_Usage_C_Value != 0 && Rank_4_Usage_C != 0)
                    {
                        Site_sheet.Cells[9, 6] = Rank_4_Usage_C_Value;
                    }
                    if (CQI_C_Value != 0 && CQI_C != 0)
                    {
                        Site_sheet.Cells[10, 6] = CQI_C_Value;
                    }
                    if (MCS_C_Value != 0 && MCS_C != 0)
                    {
                        Site_sheet.Cells[11, 6] = MCS_C_Value;
                    }
                    if (DL_User_Thp_Mbps_Daily_C_Value != 0 && DL_User_Thp_Mbps_Daily_C != 0)
                    {
                        Site_sheet.Cells[13, 6] = DL_User_Thp_Mbps_Daily_C_Value;
                    }
                    //Site_sheet.Cells[13, 6] = DL_User_Thp_Mbps_23_C_Value;
                    if (UL_User_Thp_Mbps_Daily_C_Value != 0 && UL_User_Thp_Mbps_Daily_C != 0)
                    {
                        Site_sheet.Cells[15, 6] = UL_User_Thp_Mbps_Daily_C_Value;
                    }
                    //Site_sheet.Cells[15, 6] = UL_User_Thp_Mbps_23_C_Value;
                    Site_sheet.Cells[16, 6] = "";
                    if (RB_Utilizing_Rate_DL_C_Value != 0 && RB_Utilizing_Rate_DL_C != 0)
                    {
                        Site_sheet.Cells[17, 6] = RB_Utilizing_Rate_DL_C_Value;
                    }
                    if (RB_Utilizing_Rate_UL_C_Value != 0 && RB_Utilizing_Rate_UL_C != 0)
                    {
                        Site_sheet.Cells[18, 6] = RB_Utilizing_Rate_UL_C_Value;
                    }
                    if (Cell_Unavailable_Rate_C_Value != 0 && Cell_Unavailable_Rate_C != 0)
                    {
                        Site_sheet.Cells[19, 6] = 100 - Cell_Unavailable_Rate_C_Value ;
                    }
                    if (Average_User_Number_C_Value != 0 && Average_User_Number_C != 0)
                    {
                        Site_sheet.Cells[20, 6] = Average_User_Number_C_Value;
                    }



                    SCG_Add_D_Value = SCG_Add_D_Value / SCG_Add_D;
                    SCG_Change_D_Value = SCG_Change_D_Value / SCG_Change_D;
                    SCG_Drop_Rate_D_Value = SCG_Drop_Rate_D_Value / SCG_Drop_Rate_D;
                    //DL_traffic_GB = DL_traffic_GB ;
                    //UL_traffic_GB = UL_traffic_GB ;
                    Rank_2_Usage_D_Value = Rank_2_Usage_D_Value / Rank_2_Usage_D;
                    Rank_3_Usage_D_Value = Rank_3_Usage_D_Value / Rank_3_Usage_D;
                    Rank_4_Usage_D_Value = Rank_4_Usage_D_Value / Rank_4_Usage_D;
                    CQI_D_Value = CQI_D_Value / CQI_D;
                    MCS_D_Value = MCS_D_Value / MCS_D;
                    DL_User_Thp_Mbps_Daily_D_Value = DL_User_Thp_Mbps_Daily_D_Value / DL_User_Thp_Mbps_Daily_D;
                    //DL_User_Thp_Mbps_23_D_Value = DL_User_Thp_Mbps_23_D_Value / DL_User_Thp_Mbps_23_D;
                    UL_User_Thp_Mbps_Daily_D_Value = UL_User_Thp_Mbps_Daily_D_Value / UL_User_Thp_Mbps_Daily_D;
                    //UL_User_Thp_Mbps_23_D_Value = UL_User_Thp_Mbps_23_D_Value / UL_User_Thp_Mbps_23_D;
                    Pcell_change_Succes_Rate_D_Value = Pcell_change_Succes_Rate_D_Value / Pcell_change_Succes_Rate_D;
                    RB_Utilizing_Rate_DL_D_Value = RB_Utilizing_Rate_DL_D_Value / RB_Utilizing_Rate_DL_D;
                    RB_Utilizing_Rate_UL_D_Value = RB_Utilizing_Rate_UL_D_Value / RB_Utilizing_Rate_UL_D;
                    Cell_Unavailable_Rate_D_Value = Cell_Unavailable_Rate_D_Value / Cell_Unavailable_Rate_D;
                    Average_User_Number_D_Value = Average_User_Number_D_Value / Average_User_Number_D;



                    Site_sheet.Cells[1, 7] = "Sector D";
                    if (SCG_Add_D_Value != 0 && SCG_Add_D != 0)
                    {
                        Site_sheet.Cells[2, 7] = SCG_Add_D_Value;
                    }
                    Site_sheet.Cells[3, 7] = "";
                    if (SCG_Drop_Rate_D_Value != 0 && SCG_Drop_Rate_D != 0)
                    {
                        Site_sheet.Cells[4, 7] = SCG_Drop_Rate_D_Value;
                    }
                    if (DL_traffic_GB_D_Value != 0 && DL_traffic_GB_D != 0)
                    {
                        Site_sheet.Cells[5, 7] = DL_traffic_GB_D_Value;
                    }
                    if (UL_traffic_GB_D_Value != 0 && UL_traffic_GB_D != 0)
                    {
                        Site_sheet.Cells[6, 7] = UL_traffic_GB_D_Value;
                    }
                    if (Rank_2_Usage_D_Value != 0 && Rank_2_Usage_D != 0)
                    {
                        Site_sheet.Cells[7, 7] = Rank_2_Usage_D_Value;
                    }
                    if (Rank_3_Usage_D_Value != 0 && Rank_3_Usage_D != 0)
                    {
                        Site_sheet.Cells[8, 7] = Rank_3_Usage_D_Value;
                    }
                    if (Rank_4_Usage_D_Value != 0 && Rank_4_Usage_D != 0)
                    {
                        Site_sheet.Cells[9, 7] = Rank_4_Usage_D_Value;
                    }
                    if (CQI_D_Value != 0 && CQI_D != 0)
                    {
                        Site_sheet.Cells[10, 7] = CQI_D_Value;
                    }
                    if (MCS_D_Value != 0 && MCS_D != 0)
                    {
                        Site_sheet.Cells[11, 7] = MCS_D_Value;
                    }
                    if (DL_User_Thp_Mbps_Daily_D_Value != 0 && DL_User_Thp_Mbps_Daily_D != 0)
                    {
                        Site_sheet.Cells[13, 7] = DL_User_Thp_Mbps_Daily_D_Value;
                    }
                    //Site_sheet.Cells[13, 7] = DL_User_Thp_Mbps_23_D_Value;
                    if (UL_User_Thp_Mbps_Daily_D_Value != 0 && UL_User_Thp_Mbps_Daily_D != 0)
                    {
                        Site_sheet.Cells[15, 7] = UL_User_Thp_Mbps_Daily_D_Value;
                    }
                    //Site_sheet.Cells[15, 7] = UL_User_Thp_Mbps_23_D_Value;
                    Site_sheet.Cells[16, 7] = "";
                    if (RB_Utilizing_Rate_DL_D_Value != 0 && RB_Utilizing_Rate_DL_D != 0)
                    {
                        Site_sheet.Cells[17, 7] = RB_Utilizing_Rate_DL_D_Value;
                    }
                    if (RB_Utilizing_Rate_UL_D_Value != 0 && RB_Utilizing_Rate_UL_D != 0)
                    {
                        Site_sheet.Cells[18, 7] = RB_Utilizing_Rate_UL_D_Value;
                    }
                    if (Cell_Unavailable_Rate_D_Value != 0 && Cell_Unavailable_Rate_D != 0)
                    {
                        Site_sheet.Cells[19, 7] = 100 - Cell_Unavailable_Rate_D_Value;
                    }
                    if (Average_User_Number_D_Value != 0 && Average_User_Number_D != 0)
                    {
                        Site_sheet.Cells[20, 7] = Average_User_Number_D_Value;
                    }










                    // BH
                    for (int k2 = 0; k2 < Count2 - 1; k2++)
                    {

                        string Cell = FARAZ_Data2[k2 + 1, 2].ToString();
                        string Site1 = Cell.Substring(5, 2) + Cell.Substring(9, 4);

                        if (Site1 == Site)
                        {

                            string Sector = Cell.Substring(13, 1);

                            if (Sector == "A")
                            {

                                if (FARAZ_Data2[k2 + 1, 3] != null )
                                {
                                    DL_User_Thp_Mbps_23_A++;
                                    DL_User_Thp_Mbps_23_A_Value = DL_User_Thp_Mbps_23_A_Value + Convert.ToDouble(FARAZ_Data2[k2 + 1, 3].ToString());
                                }

                                if (FARAZ_Data2[k2 + 1, 4] != null )
                                {
                                    UL_User_Thp_Mbps_23_A++;
                                    UL_User_Thp_Mbps_23_A_Value = UL_User_Thp_Mbps_23_A_Value + Convert.ToDouble(FARAZ_Data2[k2 + 1, 4].ToString());
                                }

                            }





                            if (Sector == "B")
                            {

                                if (FARAZ_Data2[k2 + 1, 3] != null)
                                {
                                    DL_User_Thp_Mbps_23_B++;
                                    DL_User_Thp_Mbps_23_B_Value = DL_User_Thp_Mbps_23_B_Value + Convert.ToDouble(FARAZ_Data2[k2 + 1, 3].ToString());
                                }

                                if (FARAZ_Data2[k2 + 1, 4] != null)
                                {
                                    UL_User_Thp_Mbps_23_B++;
                                    UL_User_Thp_Mbps_23_B_Value = UL_User_Thp_Mbps_23_B_Value + Convert.ToDouble(FARAZ_Data2[k2 + 1, 4].ToString());
                                }


                            }






                            if (Sector == "C")
                            {


                                if (FARAZ_Data2[k2 + 1, 3] != null)
                                {
                                    DL_User_Thp_Mbps_23_C++;
                                    DL_User_Thp_Mbps_23_C_Value = DL_User_Thp_Mbps_23_C_Value + Convert.ToDouble(FARAZ_Data2[k2 + 1, 3].ToString());
                                }

                                if (FARAZ_Data2[k2 + 1, 4] != null)
                                {
                                    UL_User_Thp_Mbps_23_C++;
                                    UL_User_Thp_Mbps_23_C_Value = UL_User_Thp_Mbps_23_C_Value + Convert.ToDouble(FARAZ_Data2[k2 + 1, 4].ToString());
                                }


                            }




                            if (Sector == "D")
                            {


                                if (FARAZ_Data2[k2 + 1, 3] != null)
                                {
                                    DL_User_Thp_Mbps_23_D++;
                                    DL_User_Thp_Mbps_23_D_Value = DL_User_Thp_Mbps_23_D_Value + Convert.ToDouble(FARAZ_Data2[k2 + 1, 3].ToString());
                                }

                                if (FARAZ_Data2[k2 + 1, 4] != null)
                                {
                                    UL_User_Thp_Mbps_23_D++;
                                    UL_User_Thp_Mbps_23_D_Value = UL_User_Thp_Mbps_23_D_Value + Convert.ToDouble(FARAZ_Data2[k2 + 1, 4].ToString());
                                }


                            }








                        }



                    }





                    Site_sheet = workbook1.Worksheets[index_of_sheet];

                    DL_User_Thp_Mbps_23_A_Value = DL_User_Thp_Mbps_23_A_Value / DL_User_Thp_Mbps_23_A;
                    UL_User_Thp_Mbps_23_A_Value = UL_User_Thp_Mbps_23_A_Value / UL_User_Thp_Mbps_23_A;

                    if (DL_User_Thp_Mbps_23_A_Value!=0 && DL_User_Thp_Mbps_23_A!=0)
                    {
                        Site_sheet.Cells[12, 4] = DL_User_Thp_Mbps_23_A_Value;
                    }
                    if (UL_User_Thp_Mbps_23_A_Value != 0 && UL_User_Thp_Mbps_23_A != 0)
                    {
                        Site_sheet.Cells[14, 4] = UL_User_Thp_Mbps_23_A_Value;
                    }


                    DL_User_Thp_Mbps_23_B_Value = DL_User_Thp_Mbps_23_B_Value / DL_User_Thp_Mbps_23_B;
                    UL_User_Thp_Mbps_23_B_Value = UL_User_Thp_Mbps_23_B_Value / UL_User_Thp_Mbps_23_B;

                    if (DL_User_Thp_Mbps_23_B_Value != 0 && DL_User_Thp_Mbps_23_B != 0)
                    {
                        Site_sheet.Cells[12, 5] = DL_User_Thp_Mbps_23_B_Value;
                    }
                    if (UL_User_Thp_Mbps_23_B_Value != 0 && UL_User_Thp_Mbps_23_B != 0)
                    {
                        Site_sheet.Cells[14, 5] = UL_User_Thp_Mbps_23_B_Value;
                    }

                    DL_User_Thp_Mbps_23_C_Value = DL_User_Thp_Mbps_23_C_Value / DL_User_Thp_Mbps_23_C;
                    UL_User_Thp_Mbps_23_C_Value = UL_User_Thp_Mbps_23_C_Value / UL_User_Thp_Mbps_23_C;

                    if (DL_User_Thp_Mbps_23_C_Value != 0 && DL_User_Thp_Mbps_23_C != 0)
                    {
                        Site_sheet.Cells[12, 6] = DL_User_Thp_Mbps_23_C_Value;
                    }
                    if (UL_User_Thp_Mbps_23_C_Value != 0 && UL_User_Thp_Mbps_23_C != 0)
                    {
                        Site_sheet.Cells[14, 6] = UL_User_Thp_Mbps_23_C_Value;
                    }

                    DL_User_Thp_Mbps_23_D_Value = DL_User_Thp_Mbps_23_D_Value / DL_User_Thp_Mbps_23_D;
                    UL_User_Thp_Mbps_23_D_Value = UL_User_Thp_Mbps_23_D_Value / UL_User_Thp_Mbps_23_D;

                    if (DL_User_Thp_Mbps_23_D_Value != 0 && DL_User_Thp_Mbps_23_D != 0)
                    {
                        Site_sheet.Cells[12, 7] = DL_User_Thp_Mbps_23_D_Value;
                    }
                    if (UL_User_Thp_Mbps_23_D_Value != 0 && UL_User_Thp_Mbps_23_D != 0)
                    {
                        Site_sheet.Cells[14, 7] = UL_User_Thp_Mbps_23_D_Value;
                    }


                }



                workbook1.Save();
                workbook1.Close();
                xlApp1.Quit();



                MessageBox.Show("Finished");



            }


        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {


            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel File|*.xlsx";
            DialogResult result = openFileDialog1.ShowDialog();
            string File_Name = openFileDialog1.SafeFileName.ToString();
            if (result == DialogResult.OK)
            {

                string file = openFileDialog1.FileName;
                xlApp1 = new Excel.Application();
                workbook1 = xlApp1.Workbooks.Open(file);
                sheet1 = workbook1.Worksheets[1];
                sheet2 = workbook1.Worksheets[2];




                Excel.Range History_TT = sheet1.get_Range("A2", "Q" + sheet1.UsedRange.Rows.Count);
                object[,] FARAZ_Data = (object[,])History_TT.Value;
                int Count = sheet1.UsedRange.Rows.Count;

                Excel.Range History_TT2 = sheet2.get_Range("A2", "D" + sheet2.UsedRange.Rows.Count);
                object[,] FARAZ_Data2 = (object[,])History_TT2.Value;
                int Count2 = sheet2.UsedRange.Rows.Count;

                for (int k = 0; k < Count - 1; k++)
                {

                    string Cell = FARAZ_Data[k + 1, 2].ToString();
                    string Site = Cell.Substring(5, 2) + Cell.Substring(9, 4);

                    if (!Site_Vec.Contains(Site))
                    {
                        Site_Vec[Site_Vec_Ind] = Site;
                        Site_Vec_Ind++;

                        var xlNewSheet = (Excel.Worksheet)workbook1.Worksheets.Add(Type.Missing, workbook1.Worksheets[2], Type.Missing, Type.Missing);
                        xlNewSheet.Name = Site;
                    }


                }


                for (int k = 0; k < Site_Vec_Ind; k++)
                {


                    Site_sheet = workbook1.Worksheets[2 + k + 1];

                    Site_sheet.Cells[1, 1] = "Technology";
                    Site_sheet.Cells[2, 1] = "NR";
                    Site_sheet.Cells[3, 1] = "NR";
                    Site_sheet.Cells[4, 1] = "NR";
                    Site_sheet.Cells[5, 1] = "NR";
                    Site_sheet.Cells[6, 1] = "NR";
                    Site_sheet.Cells[7, 1] = "NR";
                    Site_sheet.Cells[8, 1] = "NR";
                    Site_sheet.Cells[9, 1] = "NR";
                    Site_sheet.Cells[10, 1] = "NR";
                    Site_sheet.Cells[11, 1] = "NR";
                    Site_sheet.Cells[12, 1] = "NR";
                    Site_sheet.Cells[13, 1] = "NR";
                    Site_sheet.Cells[14, 1] = "NR";
                    Site_sheet.Cells[15, 1] = "NR";
                    Site_sheet.Cells[16, 1] = "NR";
                    Site_sheet.Cells[17, 1] = "NR";
                    Site_sheet.Cells[18, 1] = "NR";
                    Site_sheet.Cells[19, 1] = "NR";
                    Site_sheet.Cells[20, 1] = "NR";

                    Site_sheet.Cells[1, 2] = "Category";
                    Site_sheet.Cells[2, 2] = "SCG Add";
                    Site_sheet.Cells[3, 2] = "SCG Change";
                    Site_sheet.Cells[4, 2] = "SCG  Drop Rate";
                    Site_sheet.Cells[5, 2] = "DL traffic GB";
                    Site_sheet.Cells[6, 2] = "UL traffic GB";
                    Site_sheet.Cells[7, 2] = "Rank 2 Usage(%)";
                    Site_sheet.Cells[8, 2] = "Rank 3 Usage(%)";
                    Site_sheet.Cells[9, 2] = "Rank 4 Usage(%)";
                    Site_sheet.Cells[10, 2] = "CQI";
                    Site_sheet.Cells[11, 2] = "MCS";
                    Site_sheet.Cells[12, 2] = "DL User Thp Mbps (FH=23)";
                    Site_sheet.Cells[13, 2] = "DL User Thp Mbps (Daily)";
                    Site_sheet.Cells[14, 2] = "UL User Thp Mbps (FH=23)";
                    Site_sheet.Cells[15, 2] = "UL User Thp Mbps (Daily)";
                    Site_sheet.Cells[16, 2] = "Pcell change Succes Rate";
                    Site_sheet.Cells[17, 2] = "RB_Utilizing Rate DL";
                    Site_sheet.Cells[18, 2] = "RB_Utilizing Rate UL";
                    Site_sheet.Cells[19, 2] = "Cell Unavailable Rate";
                    Site_sheet.Cells[20, 2] = "Average User Number";


                    //Site_sheet.Cells[1, 3] = "KPI Name ";
                    //Site_sheet.Cells[2, 3] = "EN_DC_Setup_Sucess_Rate_Captured_in_gNodeb%";
                    //Site_sheet.Cells[3, 3] = "";
                    //Site_sheet.Cells[4, 3] = "Endc_drop_rate%";
                    //Site_sheet.Cells[5, 3] = "MAC_DL_Traffic_DRB+SRB_Gbyte";
                    //Site_sheet.Cells[6, 3] = "MAC_UL_Traffic_DRB+SRB_Gbyte";
                    //Site_sheet.Cells[7, 3] = "rank_2_report%";
                    //Site_sheet.Cells[8, 3] = "rank_3_report%";
                    //Site_sheet.Cells[9, 3] = "rank_4_report%";
                    //Site_sheet.Cells[10, 3] = "Average_CQI_256QAM";
                    //Site_sheet.Cells[11, 3] = "Average_MCS_Table2_Downlink_up_to_256QAM";
                    //Site_sheet.Cells[12, 3] = "Average_Downlink_MAC_User_Throughput_Mbps";
                    //Site_sheet.Cells[13, 3] = "Average_Downlink_MAC_User_Throughput_Mbps";
                    //Site_sheet.Cells[14, 3] = "Average_Uplink_MAC_User_Throughput_Mbps";
                    //Site_sheet.Cells[15, 3] = "Average_Uplink_MAC_User_Throughput_Mbps";
                    //Site_sheet.Cells[16, 3] = "";
                    //Site_sheet.Cells[17, 3] = "downlink_resource_block_utilization%";
                    //Site_sheet.Cells[18, 3] = "uplink_resource_block_utilization%";
                    //Site_sheet.Cells[19, 3] = "100-5G_Cell_Availability_Rate%";
                    //Site_sheet.Cells[20, 3] = "Average_Of_Average_Number_of_RRC_Connected_ENDC_NSA";




                    Site_sheet.Cells[1, 3] = "Faraz Name";
                    Site_sheet.Cells[2, 3] = "Average of SgNB_Addition_Success_Rate_gNodeb_side%";
                    Site_sheet.Cells[3, 3] = "Average of SCG_Change_Success_Rate%_Lte_Side";
                    Site_sheet.Cells[4, 3] = "Average of SgNB_Triggered_SgNB_Abnormal_Release_Rate%";
                    Site_sheet.Cells[5, 3] = "Sum of Downlink_Traffic_GB_RLC_Layer";
                    Site_sheet.Cells[6, 3] = "Sum of Uplink_Traffic_GB_RLC_Layer";
                    Site_sheet.Cells[7, 3] = "Average_MCS_PDSCH_64QAM_Rank2";
                    Site_sheet.Cells[8, 3] = "Average_MCS_PDSCH_64QAM_Rank3";
                    Site_sheet.Cells[9, 3] = "Average_MCS_PDSCH_64QAM_Rank4";
                    Site_sheet.Cells[10, 3] = "Average_CQI_64QAM+256QAM";
                    Site_sheet.Cells[11, 3] = "Average_MCS_Downlink_pdsch_Huawei_5G_Cell";
                    Site_sheet.Cells[12, 3] = "Average of Downlink_User_Throughput_Mbps_RLC_Layer";
                    Site_sheet.Cells[13, 3] = "Average of Downlink_User_Throughput_Mbps_RLC_Layer";
                    Site_sheet.Cells[14, 3] = "Average of Uplink_User_Throughput_Mbps_RLC_Layer";
                    Site_sheet.Cells[15, 3] = "Average of Uplink_User_Throughput_Mbps_RLC_Layer";
                    Site_sheet.Cells[16, 3] = "Average of PCell_Change_Executions_Success_Rate%";
                    Site_sheet.Cells[17, 3] = "Average of Downlink_Resource_Block_Utilizing_Rate%";
                    Site_sheet.Cells[18, 3] = "Average of Uplink_Resource_Block_Utilizing_Rate%";
                    Site_sheet.Cells[19, 3] = "100-(Average of Cell_Availability_Rate_Huawei_5G%/24)";
                    Site_sheet.Cells[20, 3] = "Average of MAX_Of_Average_Number_of_LTE_NR_NSA_DC_UEs_Huawei_Cell_5G";



                }

                for (int n = 0; n < Site_Vec_Ind; n++)
                {

                    string Site = Site_Vec[n].ToString();
                    Count = sheet1.UsedRange.Rows.Count;
                    int index_of_sheet = Site_Vec_Ind + 2 - n;
                    Count2 = sheet2.UsedRange.Rows.Count;

                    double SCG_Add_A_Value = 0;
                    double SCG_Change_A_Value = 0;
                    double SCG_Drop_Rate_A_Value = 0;
                    double DL_traffic_GB_A_Value = 0;
                    double UL_traffic_GB_A_Value = 0;
                    double Rank_2_Usage_A_Value = 0;
                    double Rank_3_Usage_A_Value = 0;
                    double Rank_4_Usage_A_Value = 0;
                    double CQI_A_Value = 0;
                    double MCS_A_Value = 0;
                    double DL_User_Thp_Mbps_Daily_A_Value = 0;
                    double DL_User_Thp_Mbps_23_A_Value = 0;
                    double UL_User_Thp_Mbps_Daily_A_Value = 0;
                    double UL_User_Thp_Mbps_23_A_Value = 0;
                    double Pcell_change_Succes_Rate_A_Value = 0;
                    double RB_Utilizing_Rate_DL_A_Value = 0;
                    double RB_Utilizing_Rate_UL_A_Value = 0;
                    double Cell_Unavailable_Rate_A_Value = 0;
                    double Average_User_Number_A_Value = 0;




                    int SCG_Add_A = 0;
                    int SCG_Change_A = 0;
                    int SCG_Drop_Rate_A = 0;
                    int DL_traffic_GB_A = 0;
                    int UL_traffic_GB_A = 0;
                    int Rank_2_Usage_A = 0;
                    int Rank_3_Usage_A = 0;
                    int Rank_4_Usage_A = 0;
                    int CQI_A = 0;
                    int MCS_A = 0;
                    int DL_User_Thp_Mbps_Daily_A = 0;
                    int DL_User_Thp_Mbps_23_A = 0;
                    int UL_User_Thp_Mbps_Daily_A = 0;
                    int UL_User_Thp_Mbps_23_A = 0;
                    int Pcell_change_Succes_Rate_A = 0;
                    int RB_Utilizing_Rate_DL_A = 0;
                    int RB_Utilizing_Rate_UL_A = 0;
                    int Cell_Unavailable_Rate_A = 0;
                    int Average_User_Number_A = 0;

                    double SCG_Add_B_Value = 0;
                    double SCG_Change_B_Value = 0;
                    double SCG_Drop_Rate_B_Value = 0;
                    double DL_traffic_GB_B_Value = 0;
                    double UL_traffic_GB_B_Value = 0;
                    double Rank_2_Usage_B_Value = 0;
                    double Rank_3_Usage_B_Value = 0;
                    double Rank_4_Usage_B_Value = 0;
                    double CQI_B_Value = 0;
                    double MCS_B_Value = 0;
                    double DL_User_Thp_Mbps_Daily_B_Value = 0;
                    double DL_User_Thp_Mbps_23_B_Value = 0;
                    double UL_User_Thp_Mbps_Daily_B_Value = 0;
                    double UL_User_Thp_Mbps_23_B_Value = 0;
                    double Pcell_change_Succes_Rate_B_Value = 0;
                    double RB_Utilizing_Rate_DL_B_Value = 0;
                    double RB_Utilizing_Rate_UL_B_Value = 0;
                    double Cell_Unavailable_Rate_B_Value = 0;
                    double Average_User_Number_B_Value = 0;


                    int SCG_Add_B = 0;
                    int SCG_Change_B = 0;
                    int SCG_Drop_Rate_B = 0;
                    int DL_traffic_GB_B = 0;
                    int UL_traffic_GB_B = 0;
                    int Rank_2_Usage_B = 0;
                    int Rank_3_Usage_B = 0;
                    int Rank_4_Usage_B = 0;
                    int CQI_B = 0;
                    int MCS_B = 0;
                    int DL_User_Thp_Mbps_Daily_B = 0;
                    int DL_User_Thp_Mbps_23_B = 0;
                    int UL_User_Thp_Mbps_Daily_B = 0;
                    int UL_User_Thp_Mbps_23_B = 0;
                    int Pcell_change_Succes_Rate_B = 0;
                    int RB_Utilizing_Rate_DL_B = 0;
                    int RB_Utilizing_Rate_UL_B = 0;
                    int Cell_Unavailable_Rate_B = 0;
                    int Average_User_Number_B = 0;



                    double SCG_Add_C_Value = 0;
                    double SCG_Change_C_Value = 0;
                    double SCG_Drop_Rate_C_Value = 0;
                    double DL_traffic_GB_C_Value = 0;
                    double UL_traffic_GB_C_Value = 0;
                    double Rank_2_Usage_C_Value = 0;
                    double Rank_3_Usage_C_Value = 0;
                    double Rank_4_Usage_C_Value = 0;
                    double CQI_C_Value = 0;
                    double MCS_C_Value = 0;
                    double DL_User_Thp_Mbps_Daily_C_Value = 0;
                    double DL_User_Thp_Mbps_23_C_Value = 0;
                    double UL_User_Thp_Mbps_Daily_C_Value = 0;
                    double UL_User_Thp_Mbps_23_C_Value = 0;
                    double Pcell_change_Succes_Rate_C_Value = 0;
                    double RB_Utilizing_Rate_DL_C_Value = 0;
                    double RB_Utilizing_Rate_UL_C_Value = 0;
                    double Cell_Unavailable_Rate_C_Value = 0;
                    double Average_User_Number_C_Value = 0;


                    int SCG_Add_C = 0;
                    int SCG_Change_C = 0;
                    int SCG_Drop_Rate_C = 0;
                    int DL_traffic_GB_C = 0;
                    int UL_traffic_GB_C = 0;
                    int Rank_2_Usage_C = 0;
                    int Rank_3_Usage_C = 0;
                    int Rank_4_Usage_C = 0;
                    int CQI_C = 0;
                    int MCS_C = 0;
                    int DL_User_Thp_Mbps_Daily_C = 0;
                    int DL_User_Thp_Mbps_23_C = 0;
                    int UL_User_Thp_Mbps_Daily_C = 0;
                    int UL_User_Thp_Mbps_23_C = 0;
                    int Pcell_change_Succes_Rate_C = 0;
                    int RB_Utilizing_Rate_DL_C = 0;
                    int RB_Utilizing_Rate_UL_C = 0;
                    int Cell_Unavailable_Rate_C = 0;
                    int Average_User_Number_C = 0;




                    double SCG_Add_D_Value = 0;
                    double SCG_Change_D_Value = 0;
                    double SCG_Drop_Rate_D_Value = 0;
                    double DL_traffic_GB_D_Value = 0;
                    double UL_traffic_GB_D_Value = 0;
                    double Rank_2_Usage_D_Value = 0;
                    double Rank_3_Usage_D_Value = 0;
                    double Rank_4_Usage_D_Value = 0;
                    double CQI_D_Value = 0;
                    double MCS_D_Value = 0;
                    double DL_User_Thp_Mbps_Daily_D_Value = 0;
                    double DL_User_Thp_Mbps_23_D_Value = 0;
                    double UL_User_Thp_Mbps_Daily_D_Value = 0;
                    double UL_User_Thp_Mbps_23_D_Value = 0;
                    double Pcell_change_Succes_Rate_D_Value = 0;
                    double RB_Utilizing_Rate_DL_D_Value = 0;
                    double RB_Utilizing_Rate_UL_D_Value = 0;
                    double Cell_Unavailable_Rate_D_Value = 0;
                    double Average_User_Number_D_Value = 0;


                    int SCG_Add_D = 0;
                    int SCG_Change_D = 0;
                    int SCG_Drop_Rate_D = 0;
                    int DL_traffic_GB_D = 0;
                    int UL_traffic_GB_D = 0;
                    int Rank_2_Usage_D = 0;
                    int Rank_3_Usage_D = 0;
                    int Rank_4_Usage_D = 0;
                    int CQI_D = 0;
                    int MCS_D = 0;
                    int DL_User_Thp_Mbps_Daily_D = 0;
                    int DL_User_Thp_Mbps_23_D = 0;
                    int UL_User_Thp_Mbps_Daily_D = 0;
                    int UL_User_Thp_Mbps_23_D = 0;
                    int Pcell_change_Succes_Rate_D = 0;
                    int RB_Utilizing_Rate_DL_D = 0;
                    int RB_Utilizing_Rate_UL_D = 0;
                    int Cell_Unavailable_Rate_D = 0;
                    int Average_User_Number_D = 0;



                    for (int k = 0; k < Count - 1; k++)
                    {

                        string Cell = FARAZ_Data[k + 1, 2].ToString();
                        string Site1 = Cell.Substring(5, 2) + Cell.Substring(9, 4);

                        if (Site1 == Site)
                        {

                            string Sector = Cell.Substring(13, 1);

                            if (Sector == "A")
                            {
                                if (FARAZ_Data[k + 1, 3] != null)
                                {
                                    SCG_Add_A++;
                                    SCG_Add_A_Value = SCG_Add_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 3].ToString());
                                }

                                //string SCG_Change_str = FARAZ_Data[k + 1, 4].ToString();
                                //if (SCG_Change_str != null && SCG_Change_str != "")
                                //{
                                //    SCG_Change_A++;
                                //    SCG_Change = SCG_Change + Convert.ToDouble(FARAZ_Data[k + 1, 4].ToString());
                                //}

                                if (FARAZ_Data[k + 1, 4] != null)
                                {
                                    SCG_Drop_Rate_A++;
                                    SCG_Drop_Rate_A_Value = SCG_Drop_Rate_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 4].ToString());
                                }


                                if (FARAZ_Data[k + 1, 5] != null)
                                {
                                    DL_traffic_GB_A++;
                                    DL_traffic_GB_A_Value = DL_traffic_GB_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 5].ToString());
                                }



                                if (FARAZ_Data[k + 1, 6] != null)
                                {
                                    UL_traffic_GB_A++;
                                    UL_traffic_GB_A_Value = UL_traffic_GB_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 6].ToString());
                                }


                                if (FARAZ_Data[k + 1, 7] != null)
                                {
                                    Rank_2_Usage_A++;
                                    Rank_2_Usage_A_Value = Rank_2_Usage_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 7].ToString());
                                }



                                if (FARAZ_Data[k + 1, 8] != null)
                                {
                                    Rank_3_Usage_A++;
                                    Rank_3_Usage_A_Value = Rank_3_Usage_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 8].ToString());
                                }


                                if (FARAZ_Data[k + 1, 9] != null)
                                {
                                    Rank_4_Usage_A++;
                                    Rank_4_Usage_A_Value = Rank_4_Usage_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 9].ToString());
                                }

                                if (FARAZ_Data[k + 1, 10] != null)
                                {
                                    CQI_A++;
                                    CQI_A_Value = CQI_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 10].ToString());
                                }


                                if (FARAZ_Data[k + 1, 11] != null)
                                {
                                    MCS_A++;
                                    MCS_A_Value = MCS_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 11].ToString());
                                }

                                if (FARAZ_Data[k + 1, 12] != null)
                                {
                                    DL_User_Thp_Mbps_Daily_A++;
                                    DL_User_Thp_Mbps_Daily_A_Value = DL_User_Thp_Mbps_Daily_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 12].ToString());
                                }


                                //string DL_User_Thp_Mbps_23_str = FARAZ_Data[k + 1, 13].ToString();
                                //if (DL_User_Thp_Mbps_23_str != null && DL_User_Thp_Mbps_23_str != "")
                                //{
                                //    DL_User_Thp_Mbps_23_A++;
                                //    DL_User_Thp_Mbps_23 = DL_User_Thp_Mbps_23 + Convert.ToDouble(FARAZ_Data[k + 1, 13].ToString());
                                //}

                                if (FARAZ_Data[k + 1, 13] != null)
                                {
                                    UL_User_Thp_Mbps_Daily_A++;
                                    UL_User_Thp_Mbps_Daily_A_Value = UL_User_Thp_Mbps_Daily_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 13].ToString());
                                }


                                //string UL_User_Thp_Mbps_23_str = FARAZ_Data[k + 1, 15].ToString();
                                //if (UL_User_Thp_Mbps_23_str != null && UL_User_Thp_Mbps_23_str != "")
                                //{
                                //    UL_User_Thp_Mbps_23_A++;
                                //    UL_User_Thp_Mbps_23 = UL_User_Thp_Mbps_23 + Convert.ToDouble(FARAZ_Data[k + 1, 15].ToString());
                                //}

                                //string Pcell_change_Succes_Rate_str = FARAZ_Data[k + 1, 14].ToString();
                                //if (Pcell_change_Succes_Rate_str != null && Pcell_change_Succes_Rate_str != "")
                                //{
                                //    Pcell_change_Succes_Rate_A++;
                                //    Pcell_change_Succes_Rate = Pcell_change_Succes_Rate + Convert.ToDouble(FARAZ_Data[k + 1, 14].ToString());
                                //}


                                if (FARAZ_Data[k + 1, 14] != null)
                                {
                                    RB_Utilizing_Rate_DL_A++;
                                    RB_Utilizing_Rate_DL_A_Value = RB_Utilizing_Rate_DL_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 14].ToString());
                                }


                                if (FARAZ_Data[k + 1, 15] != null)
                                {
                                    RB_Utilizing_Rate_UL_A++;
                                    RB_Utilizing_Rate_UL_A_Value = RB_Utilizing_Rate_UL_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 15].ToString());
                                }



                                if (FARAZ_Data[k + 1, 16] != null)
                                {
                                    Cell_Unavailable_Rate_A++;
                                    Cell_Unavailable_Rate_A_Value = Cell_Unavailable_Rate_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 16].ToString());
                                }



                                if (FARAZ_Data[k + 1, 17] != null)
                                {
                                    Average_User_Number_A++;
                                    Average_User_Number_A_Value = Average_User_Number_A_Value + Convert.ToDouble(FARAZ_Data[k + 1, 17].ToString());
                                }


                            }





                            if (Sector == "B")
                            {
                                if (FARAZ_Data[k + 1, 3] != null)
                                {
                                    SCG_Add_B++;
                                    SCG_Add_B_Value = SCG_Add_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 3].ToString());
                                }

                                //string SCG_Change_str = FARAZ_Data[k + 1, 4].ToString();
                                //if (SCG_Change_str != null && SCG_Change_str != "")
                                //{
                                //    SCG_Change_A++;
                                //    SCG_Change = SCG_Change + Convert.ToDouble(FARAZ_Data[k + 1, 4].ToString());
                                //}

                                if (FARAZ_Data[k + 1, 4] != null)
                                {
                                    SCG_Drop_Rate_B++;
                                    SCG_Drop_Rate_B_Value = SCG_Drop_Rate_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 4].ToString());
                                }


                                if (FARAZ_Data[k + 1, 5] != null)
                                {
                                    DL_traffic_GB_B++;
                                    DL_traffic_GB_B_Value = DL_traffic_GB_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 5].ToString());
                                }



                                if (FARAZ_Data[k + 1, 6] != null)
                                {
                                    UL_traffic_GB_B++;
                                    UL_traffic_GB_B_Value = UL_traffic_GB_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 6].ToString());
                                }


                                if (FARAZ_Data[k + 1, 7] != null)
                                {
                                    Rank_2_Usage_B++;
                                    Rank_2_Usage_B_Value = Rank_2_Usage_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 7].ToString());
                                }



                                if (FARAZ_Data[k + 1, 8] != null)
                                {
                                    Rank_3_Usage_B++;
                                    Rank_3_Usage_B_Value = Rank_3_Usage_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 8].ToString());
                                }


                                if (FARAZ_Data[k + 1, 9] != null)
                                {
                                    Rank_4_Usage_B++;
                                    Rank_4_Usage_B_Value = Rank_4_Usage_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 9].ToString());
                                }

                                if (FARAZ_Data[k + 1, 10] != null)
                                {
                                    CQI_B++;
                                    CQI_B_Value = CQI_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 10].ToString());
                                }


                                if (FARAZ_Data[k + 1, 11] != null)
                                {
                                    MCS_B++;
                                    MCS_B_Value = MCS_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 11].ToString());
                                }

                                if (FARAZ_Data[k + 1, 12] != null)
                                {
                                    DL_User_Thp_Mbps_Daily_B++;
                                    DL_User_Thp_Mbps_Daily_B_Value = DL_User_Thp_Mbps_Daily_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 12].ToString());
                                }


                                //string DL_User_Thp_Mbps_23_str = FARAZ_Data[k + 1, 13].ToString();
                                //if (DL_User_Thp_Mbps_23_str != null && DL_User_Thp_Mbps_23_str != "")
                                //{
                                //    DL_User_Thp_Mbps_23_A++;
                                //    DL_User_Thp_Mbps_23 = DL_User_Thp_Mbps_23 + Convert.ToDouble(FARAZ_Data[k + 1, 13].ToString());
                                //}

                                if (FARAZ_Data[k + 1, 13] != null)
                                {
                                    UL_User_Thp_Mbps_Daily_B++;
                                    UL_User_Thp_Mbps_Daily_B_Value = UL_User_Thp_Mbps_Daily_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 13].ToString());
                                }


                                //string UL_User_Thp_Mbps_23_str = FARAZ_Data[k + 1, 15].ToString();
                                //if (UL_User_Thp_Mbps_23_str != null && UL_User_Thp_Mbps_23_str != "")
                                //{
                                //    UL_User_Thp_Mbps_23_A++;
                                //    UL_User_Thp_Mbps_23 = UL_User_Thp_Mbps_23 + Convert.ToDouble(FARAZ_Data[k + 1, 15].ToString());
                                //}

                                //string Pcell_change_Succes_Rate_str = FARAZ_Data[k + 1, 14].ToString();
                                //if (Pcell_change_Succes_Rate_str != null && Pcell_change_Succes_Rate_str != "")
                                //{
                                //    Pcell_change_Succes_Rate_A++;
                                //    Pcell_change_Succes_Rate = Pcell_change_Succes_Rate + Convert.ToDouble(FARAZ_Data[k + 1, 14].ToString());
                                //}


                                if (FARAZ_Data[k + 1, 14] != null)
                                {
                                    RB_Utilizing_Rate_DL_B++;
                                    RB_Utilizing_Rate_DL_B_Value = RB_Utilizing_Rate_DL_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 14].ToString());
                                }


                                if (FARAZ_Data[k + 1, 15] != null)
                                {
                                    RB_Utilizing_Rate_UL_B++;
                                    RB_Utilizing_Rate_UL_B_Value = RB_Utilizing_Rate_UL_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 15].ToString());
                                }



                                if (FARAZ_Data[k + 1, 16] != null)
                                {
                                    Cell_Unavailable_Rate_B++;
                                    Cell_Unavailable_Rate_B_Value = Cell_Unavailable_Rate_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 16].ToString());
                                }



                                if (FARAZ_Data[k + 1, 17] != null)
                                {
                                    Average_User_Number_B++;
                                    Average_User_Number_B_Value = Average_User_Number_B_Value + Convert.ToDouble(FARAZ_Data[k + 1, 17].ToString());
                                }


                            }






                            if (Sector == "C")
                            {
                                if (FARAZ_Data[k + 1, 3] != null)
                                {
                                    SCG_Add_C++;
                                    SCG_Add_C_Value = SCG_Add_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 3].ToString());
                                }

                                //string SCG_Change_str = FARAZ_Data[k + 1, 4].ToString();
                                //if (SCG_Change_str != null && SCG_Change_str != "")
                                //{
                                //    SCG_Change_A++;
                                //    SCG_Change = SCG_Change + Convert.ToDouble(FARAZ_Data[k + 1, 4].ToString());
                                //}

                                if (FARAZ_Data[k + 1, 4] != null)
                                {
                                    SCG_Drop_Rate_C++;
                                    SCG_Drop_Rate_C_Value = SCG_Drop_Rate_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 4].ToString());
                                }


                                if (FARAZ_Data[k + 1, 5] != null)
                                {
                                    DL_traffic_GB_C++;
                                    DL_traffic_GB_C_Value = DL_traffic_GB_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 5].ToString());
                                }



                                if (FARAZ_Data[k + 1, 6] != null)
                                {
                                    UL_traffic_GB_C++;
                                    UL_traffic_GB_C_Value = UL_traffic_GB_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 6].ToString());
                                }


                                if (FARAZ_Data[k + 1, 7] != null)
                                {
                                    Rank_2_Usage_C++;
                                    Rank_2_Usage_C_Value = Rank_2_Usage_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 7].ToString());
                                }



                                if (FARAZ_Data[k + 1, 8] != null)
                                {
                                    Rank_3_Usage_C++;
                                    Rank_3_Usage_C_Value = Rank_3_Usage_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 8].ToString());
                                }


                                if (FARAZ_Data[k + 1, 9] != null)
                                {
                                    Rank_4_Usage_C++;
                                    Rank_4_Usage_C_Value = Rank_4_Usage_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 9].ToString());
                                }

                                if (FARAZ_Data[k + 1, 10] != null)
                                {
                                    CQI_C++;
                                    CQI_C_Value = CQI_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 10].ToString());
                                }


                                if (FARAZ_Data[k + 1, 11] != null)
                                {
                                    MCS_C++;
                                    MCS_C_Value = MCS_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 11].ToString());
                                }

                                if (FARAZ_Data[k + 1, 12] != null)
                                {
                                    DL_User_Thp_Mbps_Daily_C++;
                                    DL_User_Thp_Mbps_Daily_C_Value = DL_User_Thp_Mbps_Daily_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 12].ToString());
                                }


                                //string DL_User_Thp_Mbps_23_str = FARAZ_Data[k + 1, 13].ToString();
                                //if (DL_User_Thp_Mbps_23_str != null && DL_User_Thp_Mbps_23_str != "")
                                //{
                                //    DL_User_Thp_Mbps_23_A++;
                                //    DL_User_Thp_Mbps_23 = DL_User_Thp_Mbps_23 + Convert.ToDouble(FARAZ_Data[k + 1, 13].ToString());
                                //}

                                if (FARAZ_Data[k + 1, 13] != null)
                                {
                                    UL_User_Thp_Mbps_Daily_C++;
                                    UL_User_Thp_Mbps_Daily_C_Value = UL_User_Thp_Mbps_Daily_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 13].ToString());
                                }


                                //string UL_User_Thp_Mbps_23_str = FARAZ_Data[k + 1, 15].ToString();
                                //if (UL_User_Thp_Mbps_23_str != null && UL_User_Thp_Mbps_23_str != "")
                                //{
                                //    UL_User_Thp_Mbps_23_A++;
                                //    UL_User_Thp_Mbps_23 = UL_User_Thp_Mbps_23 + Convert.ToDouble(FARAZ_Data[k + 1, 15].ToString());
                                //}

                                //string Pcell_change_Succes_Rate_str = FARAZ_Data[k + 1, 14].ToString();
                                //if (Pcell_change_Succes_Rate_str != null && Pcell_change_Succes_Rate_str != "")
                                //{
                                //    Pcell_change_Succes_Rate_A++;
                                //    Pcell_change_Succes_Rate = Pcell_change_Succes_Rate + Convert.ToDouble(FARAZ_Data[k + 1, 14].ToString());
                                //}


                                if (FARAZ_Data[k + 1, 14] != null)
                                {
                                    RB_Utilizing_Rate_DL_C++;
                                    RB_Utilizing_Rate_DL_C_Value = RB_Utilizing_Rate_DL_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 14].ToString());
                                }


                                if (FARAZ_Data[k + 1, 15] != null)
                                {
                                    RB_Utilizing_Rate_UL_C++;
                                    RB_Utilizing_Rate_UL_C_Value = RB_Utilizing_Rate_UL_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 15].ToString());
                                }



                                if (FARAZ_Data[k + 1, 16] != null)
                                {
                                    Cell_Unavailable_Rate_C++;
                                    Cell_Unavailable_Rate_C_Value = Cell_Unavailable_Rate_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 16].ToString());
                                }



                                if (FARAZ_Data[k + 1, 17] != null)
                                {
                                    Average_User_Number_C++;
                                    Average_User_Number_C_Value = Average_User_Number_C_Value + Convert.ToDouble(FARAZ_Data[k + 1, 17].ToString());
                                }


                            }




                            if (Sector == "D")
                            {
                                if (FARAZ_Data[k + 1, 3] != null)
                                {
                                    SCG_Add_D++;
                                    SCG_Add_D_Value = SCG_Add_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 3].ToString());
                                }

                                //string SCG_Dhange_str = FARAZ_Data[k + 1, 4].ToString();
                                //if (SCG_Dhange_str != null && SCG_Dhange_str != "")
                                //{
                                //    SCG_Dhange_A++;
                                //    SCG_Dhange = SCG_Dhange + Convert.ToDouble(FARAZ_Data[k + 1, 4].ToString());
                                //}

                                if (FARAZ_Data[k + 1, 4] != null)
                                {
                                    SCG_Drop_Rate_D++;
                                    SCG_Drop_Rate_D_Value = SCG_Drop_Rate_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 4].ToString());
                                }


                                if (FARAZ_Data[k + 1, 5] != null)
                                {
                                    DL_traffic_GB_D++;
                                    DL_traffic_GB_D_Value = DL_traffic_GB_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 5].ToString());
                                }



                                if (FARAZ_Data[k + 1, 6] != null)
                                {
                                    UL_traffic_GB_D++;
                                    UL_traffic_GB_D_Value = UL_traffic_GB_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 6].ToString());
                                }


                                if (FARAZ_Data[k + 1, 7] != null)
                                {
                                    Rank_2_Usage_D++;
                                    Rank_2_Usage_D_Value = Rank_2_Usage_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 7].ToString());
                                }



                                if (FARAZ_Data[k + 1, 8] != null)
                                {
                                    Rank_3_Usage_D++;
                                    Rank_3_Usage_D_Value = Rank_3_Usage_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 8].ToString());
                                }


                                if (FARAZ_Data[k + 1, 9] != null)
                                {
                                    Rank_4_Usage_D++;
                                    Rank_4_Usage_D_Value = Rank_4_Usage_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 9].ToString());
                                }

                                if (FARAZ_Data[k + 1, 10] != null)
                                {
                                    CQI_D++;
                                    CQI_D_Value = CQI_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 10].ToString());
                                }


                                if (FARAZ_Data[k + 1, 11] != null)
                                {
                                    MCS_D++;
                                    MCS_D_Value = MCS_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 11].ToString());
                                }

                                if (FARAZ_Data[k + 1, 12] != null)
                                {
                                    DL_User_Thp_Mbps_Daily_D++;
                                    DL_User_Thp_Mbps_Daily_D_Value = DL_User_Thp_Mbps_Daily_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 12].ToString());
                                }


                                //string DL_User_Thp_Mbps_23_str = FARAZ_Data[k + 1, 13].ToString();
                                //if (DL_User_Thp_Mbps_23_str != null && DL_User_Thp_Mbps_23_str != "")
                                //{
                                //    DL_User_Thp_Mbps_23_A++;
                                //    DL_User_Thp_Mbps_23 = DL_User_Thp_Mbps_23 + Convert.ToDouble(FARAZ_Data[k + 1, 13].ToString());
                                //}

                                if (FARAZ_Data[k + 1, 13] != null)
                                {
                                    UL_User_Thp_Mbps_Daily_D++;
                                    UL_User_Thp_Mbps_Daily_D_Value = UL_User_Thp_Mbps_Daily_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 13].ToString());
                                }


                                //string UL_User_Thp_Mbps_23_str = FARAZ_Data[k + 1, 15].ToString();
                                //if (UL_User_Thp_Mbps_23_str != null && UL_User_Thp_Mbps_23_str != "")
                                //{
                                //    UL_User_Thp_Mbps_23_A++;
                                //    UL_User_Thp_Mbps_23 = UL_User_Thp_Mbps_23 + Convert.ToDouble(FARAZ_Data[k + 1, 15].ToString());
                                //}

                                //string Pcell_Dhange_Succes_Rate_str = FARAZ_Data[k + 1, 14].ToString();
                                //if (Pcell_Dhange_Succes_Rate_str != null && Pcell_Dhange_Succes_Rate_str != "")
                                //{
                                //    Pcell_Dhange_Succes_Rate_A++;
                                //    Pcell_Dhange_Succes_Rate = Pcell_Dhange_Succes_Rate + Convert.ToDouble(FARAZ_Data[k + 1, 14].ToString());
                                //}


                                if (FARAZ_Data[k + 1, 14] != null)
                                {
                                    RB_Utilizing_Rate_DL_D++;
                                    RB_Utilizing_Rate_DL_D_Value = RB_Utilizing_Rate_DL_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 14].ToString());
                                }


                                if (FARAZ_Data[k + 1, 15] != null)
                                {
                                    RB_Utilizing_Rate_UL_D++;
                                    RB_Utilizing_Rate_UL_D_Value = RB_Utilizing_Rate_UL_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 15].ToString());
                                }



                                if (FARAZ_Data[k + 1, 16] != null)
                                {
                                    Cell_Unavailable_Rate_D++;
                                    Cell_Unavailable_Rate_D_Value = Cell_Unavailable_Rate_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 16].ToString());
                                }



                                if (FARAZ_Data[k + 1, 17] != null)
                                {
                                    Average_User_Number_D++;
                                    Average_User_Number_D_Value = Average_User_Number_D_Value + Convert.ToDouble(FARAZ_Data[k + 1, 17].ToString());
                                }


                            }



                        }


                    }




                    Site_sheet = workbook1.Worksheets[index_of_sheet];


                    SCG_Add_A_Value = SCG_Add_A_Value / SCG_Add_A;
                    SCG_Change_A_Value = SCG_Change_A_Value / SCG_Change_A;
                    SCG_Drop_Rate_A_Value = SCG_Drop_Rate_A_Value / SCG_Drop_Rate_A;
                    //DL_traffic_GB = DL_traffic_GB ;
                    //UL_traffic_GB = UL_traffic_GB ;
                    Rank_2_Usage_A_Value = Rank_2_Usage_A_Value / Rank_2_Usage_A;
                    Rank_3_Usage_A_Value = Rank_3_Usage_A_Value / Rank_3_Usage_A;
                    Rank_4_Usage_A_Value = Rank_4_Usage_A_Value / Rank_4_Usage_A;
                    CQI_A_Value = CQI_A_Value / CQI_A;
                    MCS_A_Value = MCS_A_Value / MCS_A;
                    DL_User_Thp_Mbps_Daily_A_Value = DL_User_Thp_Mbps_Daily_A_Value / DL_User_Thp_Mbps_Daily_A;
                    //DL_User_Thp_Mbps_23_A_Value = DL_User_Thp_Mbps_23_A_Value / DL_User_Thp_Mbps_23_A;
                    UL_User_Thp_Mbps_Daily_A_Value = UL_User_Thp_Mbps_Daily_A_Value / UL_User_Thp_Mbps_Daily_A;
                    //UL_User_Thp_Mbps_23_A_Value = UL_User_Thp_Mbps_23_A_Value / UL_User_Thp_Mbps_23_A;
                    Pcell_change_Succes_Rate_A_Value = Pcell_change_Succes_Rate_A_Value / Pcell_change_Succes_Rate_A;
                    RB_Utilizing_Rate_DL_A_Value = RB_Utilizing_Rate_DL_A_Value / RB_Utilizing_Rate_DL_A;
                    RB_Utilizing_Rate_UL_A_Value = RB_Utilizing_Rate_UL_A_Value / RB_Utilizing_Rate_UL_A;
                    Cell_Unavailable_Rate_A_Value = Cell_Unavailable_Rate_A_Value / Cell_Unavailable_Rate_A;
                    Average_User_Number_A_Value = Average_User_Number_A_Value / Average_User_Number_A;





                    Site_sheet.Cells[1, 4] = "Sector A";
                    if (SCG_Add_A_Value != 0 && SCG_Add_A != 0)
                    {
                        Site_sheet.Cells[2, 4] = SCG_Add_A_Value;
                    }
                    Site_sheet.Cells[3, 4] = "";
                    if (SCG_Drop_Rate_A_Value != 0 && SCG_Drop_Rate_A != 0)
                    {
                        Site_sheet.Cells[4, 4] = SCG_Drop_Rate_A_Value;
                    }
                    if (DL_traffic_GB_A_Value != 0 && DL_traffic_GB_A != 0)
                    {
                        Site_sheet.Cells[5, 4] = DL_traffic_GB_A_Value;
                    }
                    if (UL_traffic_GB_A_Value != 0 && UL_traffic_GB_A != 0)
                    {
                        Site_sheet.Cells[6, 4] = UL_traffic_GB_A_Value;
                    }
                    if (Rank_2_Usage_A_Value != 0 && Rank_2_Usage_A != 0)
                    {
                        Site_sheet.Cells[7, 4] = Rank_2_Usage_A_Value;
                    }
                    if (Rank_3_Usage_A_Value != 0 && Rank_3_Usage_A != 0)
                    {
                        Site_sheet.Cells[8, 4] = Rank_3_Usage_A_Value;
                    }
                    if (Rank_4_Usage_A_Value != 0 && Rank_4_Usage_A != 0)
                    {
                        Site_sheet.Cells[9, 4] = Rank_4_Usage_A_Value;
                    }
                    if (CQI_A_Value != 0 && CQI_A != 0)
                    {
                        Site_sheet.Cells[10, 4] = CQI_A_Value;
                    }
                    if (MCS_A_Value != 0 && MCS_A != 0)
                    {
                        Site_sheet.Cells[11, 4] = MCS_A_Value;
                    }
                    if (DL_User_Thp_Mbps_Daily_A_Value != 0 && DL_User_Thp_Mbps_Daily_A != 0)
                    {
                        Site_sheet.Cells[13, 4] = DL_User_Thp_Mbps_Daily_A_Value;
                    }
                    //Site_sheet.Cells[13, 4] = DL_User_Thp_Mbps_23_A_Value;
                    if (UL_User_Thp_Mbps_Daily_A_Value != 0 && UL_User_Thp_Mbps_Daily_A != 0)
                    {
                        Site_sheet.Cells[15, 4] = UL_User_Thp_Mbps_Daily_A_Value;
                    }
                    //Site_sheet.Cells[15, 4] = UL_User_Thp_Mbps_23_A_Value;
                    Site_sheet.Cells[16, 4] = "";
                    if (RB_Utilizing_Rate_DL_A_Value != 0 && RB_Utilizing_Rate_DL_A != 0)
                    {
                        Site_sheet.Cells[17, 4] = RB_Utilizing_Rate_DL_A_Value;
                    }
                    if (RB_Utilizing_Rate_UL_A_Value != 0 && RB_Utilizing_Rate_UL_A != 0)
                    {
                        Site_sheet.Cells[18, 4] = RB_Utilizing_Rate_UL_A_Value;
                    }
                    if (Cell_Unavailable_Rate_A_Value != 0 && Cell_Unavailable_Rate_A != 0)
                    {
                        Site_sheet.Cells[19, 4] = 100 - Cell_Unavailable_Rate_A_Value/24;
                    }
                    if (Average_User_Number_A_Value != 0 && Average_User_Number_A != 0)
                    {
                        Site_sheet.Cells[20, 4] = Average_User_Number_A_Value;
                    }

                    SCG_Add_B_Value = SCG_Add_B_Value / SCG_Add_B;
                    SCG_Change_B_Value = SCG_Change_B_Value / SCG_Change_B;
                    SCG_Drop_Rate_B_Value = SCG_Drop_Rate_B_Value / SCG_Drop_Rate_B;
                    //DL_traffic_GB = DL_traffic_GB ;
                    //UL_traffic_GB = UL_traffic_GB ;
                    Rank_2_Usage_B_Value = Rank_2_Usage_B_Value / Rank_2_Usage_B;
                    Rank_3_Usage_B_Value = Rank_3_Usage_B_Value / Rank_3_Usage_B;
                    Rank_4_Usage_B_Value = Rank_4_Usage_B_Value / Rank_4_Usage_B;
                    CQI_B_Value = CQI_B_Value / CQI_B;
                    MCS_B_Value = MCS_B_Value / MCS_B;
                    DL_User_Thp_Mbps_Daily_B_Value = DL_User_Thp_Mbps_Daily_B_Value / DL_User_Thp_Mbps_Daily_B;
                    //DL_User_Thp_Mbps_23_B_Value = DL_User_Thp_Mbps_23_B_Value / DL_User_Thp_Mbps_23_B;
                    UL_User_Thp_Mbps_Daily_B_Value = UL_User_Thp_Mbps_Daily_B_Value / UL_User_Thp_Mbps_Daily_B;
                    //UL_User_Thp_Mbps_23_B_Value = UL_User_Thp_Mbps_23_B_Value / UL_User_Thp_Mbps_23_B;
                    Pcell_change_Succes_Rate_B_Value = Pcell_change_Succes_Rate_B_Value / Pcell_change_Succes_Rate_B;
                    RB_Utilizing_Rate_DL_B_Value = RB_Utilizing_Rate_DL_B_Value / RB_Utilizing_Rate_DL_B;
                    RB_Utilizing_Rate_UL_B_Value = RB_Utilizing_Rate_UL_B_Value / RB_Utilizing_Rate_UL_B;
                    Cell_Unavailable_Rate_B_Value = Cell_Unavailable_Rate_B_Value / Cell_Unavailable_Rate_B;
                    Average_User_Number_B_Value = Average_User_Number_B_Value / Average_User_Number_B;


                    Site_sheet.Cells[1, 5] = "Sector B";
                    if (SCG_Add_B_Value != 0 && SCG_Add_B != 0)
                    {
                        Site_sheet.Cells[2, 5] = SCG_Add_B_Value;
                    }
                    Site_sheet.Cells[3, 5] = "";
                    if (SCG_Drop_Rate_B_Value != 0 && SCG_Drop_Rate_B != 0)
                    {
                        Site_sheet.Cells[4, 5] = SCG_Drop_Rate_B_Value;
                    }
                    if (DL_traffic_GB_B_Value != 0 && DL_traffic_GB_B != 0)
                    {
                        Site_sheet.Cells[5, 5] = DL_traffic_GB_B_Value;
                    }
                    if (UL_traffic_GB_B_Value != 0 && UL_traffic_GB_B != 0)
                    {
                        Site_sheet.Cells[6, 5] = UL_traffic_GB_B_Value;
                    }
                    if (Rank_2_Usage_B_Value != 0 && Rank_2_Usage_B != 0)
                    {
                        Site_sheet.Cells[7, 5] = Rank_2_Usage_B_Value;
                    }
                    if (Rank_3_Usage_B_Value != 0 && Rank_3_Usage_B != 0)
                    {
                        Site_sheet.Cells[8, 5] = Rank_3_Usage_B_Value;
                    }
                    if (Rank_4_Usage_B_Value != 0 && Rank_4_Usage_B != 0)
                    {
                        Site_sheet.Cells[9, 5] = Rank_4_Usage_B_Value;
                    }
                    if (CQI_B_Value != 0 && CQI_B != 0)
                    {
                        Site_sheet.Cells[10, 5] = CQI_B_Value;
                    }
                    if (MCS_B_Value != 0 && MCS_B != 0)
                    {
                        Site_sheet.Cells[11, 5] = MCS_B_Value;
                    }
                    if (DL_User_Thp_Mbps_Daily_B_Value != 0 && DL_User_Thp_Mbps_Daily_B != 0)
                    {
                        Site_sheet.Cells[13, 5] = DL_User_Thp_Mbps_Daily_B_Value;
                    }
                    //Site_sheet.Cells[13, 5] = DL_User_Thp_Mbps_23_B_Value;
                    if (UL_User_Thp_Mbps_Daily_B_Value != 0 && UL_User_Thp_Mbps_Daily_B != 0)
                    {
                        Site_sheet.Cells[15, 5] = UL_User_Thp_Mbps_Daily_B_Value;
                    }
                    //Site_sheet.Cells[15, 5] = UL_User_Thp_Mbps_23_B_Value;
                    Site_sheet.Cells[16, 5] = "";
                    if (RB_Utilizing_Rate_DL_B_Value != 0 && RB_Utilizing_Rate_DL_B != 0)
                    {
                        Site_sheet.Cells[17, 5] = RB_Utilizing_Rate_DL_B_Value;
                    }
                    if (RB_Utilizing_Rate_UL_B_Value != 0 && RB_Utilizing_Rate_UL_B != 0)
                    {
                        Site_sheet.Cells[18, 5] = RB_Utilizing_Rate_UL_B_Value;
                    }
                    if (Cell_Unavailable_Rate_B_Value != 0 && Cell_Unavailable_Rate_B != 0)
                    {
                        Site_sheet.Cells[19, 5] = 100 - Cell_Unavailable_Rate_B_Value/24;
                    }
                    if (Average_User_Number_B_Value != 0 && Average_User_Number_B != 0)
                    {
                        Site_sheet.Cells[20, 5] = Average_User_Number_B_Value;
                    }


                    SCG_Add_C_Value = SCG_Add_C_Value / SCG_Add_C;
                    SCG_Change_C_Value = SCG_Change_C_Value / SCG_Change_C;
                    SCG_Drop_Rate_C_Value = SCG_Drop_Rate_C_Value / SCG_Drop_Rate_C;
                    //DL_traffic_GB = DL_traffic_GB ;
                    //UL_traffic_GB = UL_traffic_GB ;
                    Rank_2_Usage_C_Value = Rank_2_Usage_C_Value / Rank_2_Usage_C;
                    Rank_3_Usage_C_Value = Rank_3_Usage_C_Value / Rank_3_Usage_C;
                    Rank_4_Usage_C_Value = Rank_4_Usage_C_Value / Rank_4_Usage_C;
                    CQI_C_Value = CQI_C_Value / CQI_C;
                    MCS_C_Value = MCS_C_Value / MCS_C;
                    DL_User_Thp_Mbps_Daily_C_Value = DL_User_Thp_Mbps_Daily_C_Value / DL_User_Thp_Mbps_Daily_C;
                    //DL_User_Thp_Mbps_23_C_Value = DL_User_Thp_Mbps_23_C_Value / DL_User_Thp_Mbps_23_C;
                    UL_User_Thp_Mbps_Daily_C_Value = UL_User_Thp_Mbps_Daily_C_Value / UL_User_Thp_Mbps_Daily_C;
                    //UL_User_Thp_Mbps_23_C_Value = UL_User_Thp_Mbps_23_C_Value / UL_User_Thp_Mbps_23_C;
                    Pcell_change_Succes_Rate_C_Value = Pcell_change_Succes_Rate_C_Value / Pcell_change_Succes_Rate_C;
                    RB_Utilizing_Rate_DL_C_Value = RB_Utilizing_Rate_DL_C_Value / RB_Utilizing_Rate_DL_C;
                    RB_Utilizing_Rate_UL_C_Value = RB_Utilizing_Rate_UL_C_Value / RB_Utilizing_Rate_UL_C;
                    Cell_Unavailable_Rate_C_Value = Cell_Unavailable_Rate_C_Value / Cell_Unavailable_Rate_C;
                    Average_User_Number_C_Value = Average_User_Number_C_Value / Average_User_Number_C;




                    Site_sheet.Cells[1, 6] = "Sector C";
                    if (SCG_Add_C_Value != 0 && SCG_Add_C != 0)
                    {
                        Site_sheet.Cells[2, 6] = SCG_Add_C_Value;
                    }
                    Site_sheet.Cells[3, 6] = "";
                    if (SCG_Drop_Rate_C_Value != 0 && SCG_Drop_Rate_C != 0)
                    {
                        Site_sheet.Cells[4, 6] = SCG_Drop_Rate_C_Value;
                    }
                    if (DL_traffic_GB_C_Value != 0 && DL_traffic_GB_C != 0)
                    {
                        Site_sheet.Cells[5, 6] = DL_traffic_GB_C_Value;
                    }
                    if (UL_traffic_GB_C_Value != 0 && UL_traffic_GB_C != 0)
                    {
                        Site_sheet.Cells[6, 6] = UL_traffic_GB_C_Value;
                    }
                    if (Rank_2_Usage_C_Value != 0 && Rank_2_Usage_C != 0)
                    {
                        Site_sheet.Cells[7, 6] = Rank_2_Usage_C_Value;
                    }
                    if (Rank_3_Usage_C_Value != 0 && Rank_3_Usage_C != 0)
                    {
                        Site_sheet.Cells[8, 6] = Rank_3_Usage_C_Value;
                    }
                    if (Rank_4_Usage_C_Value != 0 && Rank_4_Usage_C != 0)
                    {
                        Site_sheet.Cells[9, 6] = Rank_4_Usage_C_Value;
                    }
                    if (CQI_C_Value != 0 && CQI_C != 0)
                    {
                        Site_sheet.Cells[10, 6] = CQI_C_Value;
                    }
                    if (MCS_C_Value != 0 && MCS_C != 0)
                    {
                        Site_sheet.Cells[11, 6] = MCS_C_Value;
                    }
                    if (DL_User_Thp_Mbps_Daily_C_Value != 0 && DL_User_Thp_Mbps_Daily_C != 0)
                    {
                        Site_sheet.Cells[13, 6] = DL_User_Thp_Mbps_Daily_C_Value;
                    }
                    //Site_sheet.Cells[13, 6] = DL_User_Thp_Mbps_23_C_Value;
                    if (UL_User_Thp_Mbps_Daily_C_Value != 0 && UL_User_Thp_Mbps_Daily_C != 0)
                    {
                        Site_sheet.Cells[15, 6] = UL_User_Thp_Mbps_Daily_C_Value;
                    }
                    //Site_sheet.Cells[15, 6] = UL_User_Thp_Mbps_23_C_Value;
                    Site_sheet.Cells[16, 6] = "";
                    if (RB_Utilizing_Rate_DL_C_Value != 0 && RB_Utilizing_Rate_DL_C != 0)
                    {
                        Site_sheet.Cells[17, 6] = RB_Utilizing_Rate_DL_C_Value;
                    }
                    if (RB_Utilizing_Rate_UL_C_Value != 0 && RB_Utilizing_Rate_UL_C != 0)
                    {
                        Site_sheet.Cells[18, 6] = RB_Utilizing_Rate_UL_C_Value;
                    }
                    if (Cell_Unavailable_Rate_C_Value != 0 && Cell_Unavailable_Rate_C != 0)
                    {
                        Site_sheet.Cells[19, 6] = 100 - Cell_Unavailable_Rate_C_Value/24;
                    }
                    if (Average_User_Number_C_Value != 0 && Average_User_Number_C != 0)
                    {
                        Site_sheet.Cells[20, 6] = Average_User_Number_C_Value;
                    }



                    SCG_Add_D_Value = SCG_Add_D_Value / SCG_Add_D;
                    SCG_Change_D_Value = SCG_Change_D_Value / SCG_Change_D;
                    SCG_Drop_Rate_D_Value = SCG_Drop_Rate_D_Value / SCG_Drop_Rate_D;
                    //DL_traffic_GB = DL_traffic_GB ;
                    //UL_traffic_GB = UL_traffic_GB ;
                    Rank_2_Usage_D_Value = Rank_2_Usage_D_Value / Rank_2_Usage_D;
                    Rank_3_Usage_D_Value = Rank_3_Usage_D_Value / Rank_3_Usage_D;
                    Rank_4_Usage_D_Value = Rank_4_Usage_D_Value / Rank_4_Usage_D;
                    CQI_D_Value = CQI_D_Value / CQI_D;
                    MCS_D_Value = MCS_D_Value / MCS_D;
                    DL_User_Thp_Mbps_Daily_D_Value = DL_User_Thp_Mbps_Daily_D_Value / DL_User_Thp_Mbps_Daily_D;
                    //DL_User_Thp_Mbps_23_D_Value = DL_User_Thp_Mbps_23_D_Value / DL_User_Thp_Mbps_23_D;
                    UL_User_Thp_Mbps_Daily_D_Value = UL_User_Thp_Mbps_Daily_D_Value / UL_User_Thp_Mbps_Daily_D;
                    //UL_User_Thp_Mbps_23_D_Value = UL_User_Thp_Mbps_23_D_Value / UL_User_Thp_Mbps_23_D;
                    Pcell_change_Succes_Rate_D_Value = Pcell_change_Succes_Rate_D_Value / Pcell_change_Succes_Rate_D;
                    RB_Utilizing_Rate_DL_D_Value = RB_Utilizing_Rate_DL_D_Value / RB_Utilizing_Rate_DL_D;
                    RB_Utilizing_Rate_UL_D_Value = RB_Utilizing_Rate_UL_D_Value / RB_Utilizing_Rate_UL_D;
                    Cell_Unavailable_Rate_D_Value = Cell_Unavailable_Rate_D_Value / Cell_Unavailable_Rate_D;
                    Average_User_Number_D_Value = Average_User_Number_D_Value / Average_User_Number_D;



                    Site_sheet.Cells[1, 7] = "Sector D";
                    if (SCG_Add_D_Value != 0 && SCG_Add_D != 0)
                    {
                        Site_sheet.Cells[2, 7] = SCG_Add_D_Value;
                    }
                    Site_sheet.Cells[3, 7] = "";
                    if (SCG_Drop_Rate_D_Value != 0 && SCG_Drop_Rate_D != 0)
                    {
                        Site_sheet.Cells[4, 7] = SCG_Drop_Rate_D_Value;
                    }
                    if (DL_traffic_GB_D_Value != 0 && DL_traffic_GB_D != 0)
                    {
                        Site_sheet.Cells[5, 7] = DL_traffic_GB_D_Value;
                    }
                    if (UL_traffic_GB_D_Value != 0 && UL_traffic_GB_D != 0)
                    {
                        Site_sheet.Cells[6, 7] = UL_traffic_GB_D_Value;
                    }
                    if (Rank_2_Usage_D_Value != 0 && Rank_2_Usage_D != 0)
                    {
                        Site_sheet.Cells[7, 7] = Rank_2_Usage_D_Value;
                    }
                    if (Rank_3_Usage_D_Value != 0 && Rank_3_Usage_D != 0)
                    {
                        Site_sheet.Cells[8, 7] = Rank_3_Usage_D_Value;
                    }
                    if (Rank_4_Usage_D_Value != 0 && Rank_4_Usage_D != 0)
                    {
                        Site_sheet.Cells[9, 7] = Rank_4_Usage_D_Value;
                    }
                    if (CQI_D_Value != 0 && CQI_D != 0)
                    {
                        Site_sheet.Cells[10, 7] = CQI_D_Value;
                    }
                    if (MCS_D_Value != 0 && MCS_D != 0)
                    {
                        Site_sheet.Cells[11, 7] = MCS_D_Value;
                    }
                    if (DL_User_Thp_Mbps_Daily_D_Value != 0 && DL_User_Thp_Mbps_Daily_D != 0)
                    {
                        Site_sheet.Cells[13, 7] = DL_User_Thp_Mbps_Daily_D_Value;
                    }
                    //Site_sheet.Cells[13, 7] = DL_User_Thp_Mbps_23_D_Value;
                    if (UL_User_Thp_Mbps_Daily_D_Value != 0 && UL_User_Thp_Mbps_Daily_D != 0)
                    {
                        Site_sheet.Cells[15, 7] = UL_User_Thp_Mbps_Daily_D_Value;
                    }
                    //Site_sheet.Cells[15, 7] = UL_User_Thp_Mbps_23_D_Value;
                    Site_sheet.Cells[16, 7] = "";
                    if (RB_Utilizing_Rate_DL_D_Value != 0 && RB_Utilizing_Rate_DL_D != 0)
                    {
                        Site_sheet.Cells[17, 7] = RB_Utilizing_Rate_DL_D_Value;
                    }
                    if (RB_Utilizing_Rate_UL_D_Value != 0 && RB_Utilizing_Rate_UL_D != 0)
                    {
                        Site_sheet.Cells[18, 7] = RB_Utilizing_Rate_UL_D_Value;
                    }
                    if (Cell_Unavailable_Rate_D_Value != 0 && Cell_Unavailable_Rate_D != 0)
                    {
                        Site_sheet.Cells[19, 7] = 100 - Cell_Unavailable_Rate_D_Value/24;
                    }
                    if (Average_User_Number_D_Value != 0 && Average_User_Number_D != 0)
                    {
                        Site_sheet.Cells[20, 7] = Average_User_Number_D_Value;
                    }









                    // BH
                    for (int k2 = 0; k2 < Count2 - 1; k2++)
                    {

                        string Cell = FARAZ_Data2[k2 + 1, 2].ToString();
                        string Site1 = Cell.Substring(5, 2) + Cell.Substring(9, 4);

                        if (Site1 == Site)
                        {

                            string Sector = Cell.Substring(13, 1);

                            if (Sector == "A")
                            {

                                if (FARAZ_Data2[k2 + 1, 3] != null)
                                {
                                    DL_User_Thp_Mbps_23_A++;
                                    DL_User_Thp_Mbps_23_A_Value = DL_User_Thp_Mbps_23_A_Value + Convert.ToDouble(FARAZ_Data2[k2 + 1, 3].ToString());
                                }

                                if (FARAZ_Data2[k2 + 1, 4] != null)
                                {
                                    UL_User_Thp_Mbps_23_A++;
                                    UL_User_Thp_Mbps_23_A_Value = UL_User_Thp_Mbps_23_A_Value + Convert.ToDouble(FARAZ_Data2[k2 + 1, 4].ToString());
                                }

                            }





                            if (Sector == "B")
                            {

                                if (FARAZ_Data2[k2 + 1, 3] != null)
                                {
                                    DL_User_Thp_Mbps_23_B++;
                                    DL_User_Thp_Mbps_23_B_Value = DL_User_Thp_Mbps_23_B_Value + Convert.ToDouble(FARAZ_Data2[k2 + 1, 3].ToString());
                                }

                                if (FARAZ_Data2[k2 + 1, 4] != null)
                                {
                                    UL_User_Thp_Mbps_23_B++;
                                    UL_User_Thp_Mbps_23_B_Value = UL_User_Thp_Mbps_23_B_Value + Convert.ToDouble(FARAZ_Data2[k2 + 1, 4].ToString());
                                }


                            }






                            if (Sector == "C")
                            {


                                if (FARAZ_Data2[k2 + 1, 3] != null)
                                {
                                    DL_User_Thp_Mbps_23_C++;
                                    DL_User_Thp_Mbps_23_C_Value = DL_User_Thp_Mbps_23_C_Value + Convert.ToDouble(FARAZ_Data2[k2 + 1, 3].ToString());
                                }

                                if (FARAZ_Data2[k2 + 1, 4] != null)
                                {
                                    UL_User_Thp_Mbps_23_C++;
                                    UL_User_Thp_Mbps_23_C_Value = UL_User_Thp_Mbps_23_C_Value + Convert.ToDouble(FARAZ_Data2[k2 + 1, 4].ToString());
                                }


                            }




                            if (Sector == "D")
                            {


                                if (FARAZ_Data2[k2 + 1, 3] != null)
                                {
                                    DL_User_Thp_Mbps_23_D++;
                                    DL_User_Thp_Mbps_23_D_Value = DL_User_Thp_Mbps_23_D_Value + Convert.ToDouble(FARAZ_Data2[k2 + 1, 3].ToString());
                                }

                                if (FARAZ_Data2[k2 + 1, 4] != null)
                                {
                                    UL_User_Thp_Mbps_23_D++;
                                    UL_User_Thp_Mbps_23_D_Value = UL_User_Thp_Mbps_23_D_Value + Convert.ToDouble(FARAZ_Data2[k2 + 1, 4].ToString());
                                }


                            }








                        }



                    }





                    Site_sheet = workbook1.Worksheets[index_of_sheet];

                    DL_User_Thp_Mbps_23_A_Value = DL_User_Thp_Mbps_23_A_Value / DL_User_Thp_Mbps_23_A;
                    UL_User_Thp_Mbps_23_A_Value = UL_User_Thp_Mbps_23_A_Value / UL_User_Thp_Mbps_23_A;

                    if (DL_User_Thp_Mbps_23_A_Value != 0 && DL_User_Thp_Mbps_23_A != 0)
                    {
                        Site_sheet.Cells[12, 4] = DL_User_Thp_Mbps_23_A_Value;
                    }
                    if (UL_User_Thp_Mbps_23_A_Value != 0 && UL_User_Thp_Mbps_23_A != 0)
                    {
                        Site_sheet.Cells[14, 4] = UL_User_Thp_Mbps_23_A_Value;
                    }


                    DL_User_Thp_Mbps_23_B_Value = DL_User_Thp_Mbps_23_B_Value / DL_User_Thp_Mbps_23_B;
                    UL_User_Thp_Mbps_23_B_Value = UL_User_Thp_Mbps_23_B_Value / UL_User_Thp_Mbps_23_B;

                    if (DL_User_Thp_Mbps_23_B_Value != 0 && DL_User_Thp_Mbps_23_B != 0)
                    {
                        Site_sheet.Cells[12, 5] = DL_User_Thp_Mbps_23_B_Value;
                    }
                    if (UL_User_Thp_Mbps_23_B_Value != 0 && UL_User_Thp_Mbps_23_B != 0)
                    {
                        Site_sheet.Cells[14, 5] = UL_User_Thp_Mbps_23_B_Value;
                    }

                    DL_User_Thp_Mbps_23_C_Value = DL_User_Thp_Mbps_23_C_Value / DL_User_Thp_Mbps_23_C;
                    UL_User_Thp_Mbps_23_C_Value = UL_User_Thp_Mbps_23_C_Value / UL_User_Thp_Mbps_23_C;

                    if (DL_User_Thp_Mbps_23_C_Value != 0 && DL_User_Thp_Mbps_23_C != 0)
                    {
                        Site_sheet.Cells[12, 6] = DL_User_Thp_Mbps_23_C_Value;
                    }
                    if (UL_User_Thp_Mbps_23_C_Value != 0 && UL_User_Thp_Mbps_23_C != 0)
                    {
                        Site_sheet.Cells[14, 6] = UL_User_Thp_Mbps_23_C_Value;
                    }

                    DL_User_Thp_Mbps_23_D_Value = DL_User_Thp_Mbps_23_D_Value / DL_User_Thp_Mbps_23_D;
                    UL_User_Thp_Mbps_23_D_Value = UL_User_Thp_Mbps_23_D_Value / UL_User_Thp_Mbps_23_D;

                    if (DL_User_Thp_Mbps_23_D_Value != 0 && DL_User_Thp_Mbps_23_D != 0)
                    {
                        Site_sheet.Cells[12, 7] = DL_User_Thp_Mbps_23_D_Value;
                    }
                    if (UL_User_Thp_Mbps_23_D_Value != 0 && UL_User_Thp_Mbps_23_D != 0)
                    {
                        Site_sheet.Cells[14, 7] = UL_User_Thp_Mbps_23_D_Value;
                    }


                }



                workbook1.Save();
                workbook1.Close();
                xlApp1.Quit();



                MessageBox.Show("Finished");



            }


        }
    }
}
