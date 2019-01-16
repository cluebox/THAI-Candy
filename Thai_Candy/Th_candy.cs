using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using SpssLib.SpssDataset;
using SpssLib.DataReader;
using System.Data.SqlClient;


namespace Thai_Candy
{
    class Th_candy
    {
        static void Main(string[] args)
        {
            int SurveyID = 600553;
            string SurveyPeriod = "2015-05-31";//Survey Period
            string Country = "THAILAND";//survey country
            ThInsertion iobj = new ThInsertion();
            string questions = "ID,Quota_Gender,E3,Quota_Location,Quota_Age,E2,E1,S4.1_BKK,S4.2_BKK,S4.1_UPC,S4.2_UPC,Wave,S5_1,S5_2,S5_3,S5_4,S5_5,S5_6,S5_7,S5_8,S5_9,S5_10,S5_11,S5_12,S5_13,S5_14,S5_15,S5_16,S5_17,S5_18,S5_19,S6,S7_1,S7_2,S7_3,S7_4,S7_5,S7_6,S7_7,S7_8,S7_9,S7_10,S7_11,S7_12,S7_13,S7_14,S7_15,S7_16,S7_17,S7_18,S7_19,S8,S10.2,S10.3,S11.2,S11.3,r37.1,r37.2_2,r37.2_6,r37.2_7,r37.2_8,r37.2_9,r37.2_26,r37.2_27,r37.2_28,r37.2_29,r37.2_30,r37.2_31,r37.2_35,r37.2_11,r37.2_3,r37.2_4,r37.2_12,r37.2_10,r39_2,r39_6,r39_7,r39_8,r39_9,r39_26,r39_27,r39_28,r39_29,r39_30,r39_31,r39_35,r39_11,r39_3,r39_4,r39_12,r39_10,r38.1,r38.2_2,r38.2_6,r38.2_7,r38.2_8,r38.2_9,r38.2_26,r38.2_27,r38.2_28,r38.2_29,r38.2_30,r38.2_31,r38.2_35,r38.2_11,r38.2_3,r38.2_4,r38.2_12,r38.2_10,r40_2,r40_6,r40_7,r40_8,r40_9,r40_26,r40_27,r40_28,r40_29,r40_30,r40_31,r40_35,r40_11,r40_3,r40_4,r40_12,r40_10,r41.1_2,r41.1_6,r41.1_7,r41.1_8,r41.1_9,r41.1_26,r41.1_27,r41.1_28,r41.1_29,r41.1_30,r41.1_31,r41.1_35,r41.1_11,r41.1_3,r41.1_4,r41.1_12,r41.1_10,r41.2_2,r41.2_6,r41.2_7,r41.2_8,r41.2_9,r41.2_26,r41.2_27,r41.2_28,r41.2_29,r41.2_30,r41.2_31,r41.2_35,r41.2_11,r41.2_3,r41.2_4,r41.2_12,r41.2_10,r42.1_2,r42.1_6,r42.1_7,r42.1_8,r42.1_9,r42.1_26,r42.1_27,r42.1_28,r42.1_29,r42.1_30,r42.1_31,r42.1_35,r42.1_11,r42.1_3,r42.1_4,r42.1_12,r42.1_10,r42.2_2,r42.2_6,r42.2_7,r42.2_8,r42.2_9,r42.2_26,r42.2_27,r42.2_28,r42.2_29,r42.2_30,r42.2_31,r42.2_35,r42.2_11,r42.2_3,r42.2_4,r42.2_12,r42.2_10,r42.4_2,r42.4_6,r42.4_7,r42.4_8,r42.4_9,r42.4_26,r42.4_27,r42.4_28,r42.4_29,r42.4_30,r42.4_31,r42.4_35,r42.4_11,r42.4_3,r42.4_4,r42.4_12,r42.4_10,r41.3,r42.3,r37.1_net,r37.2_net_1,r37.2_net_2,r37.2_net_4,r37.2_net_5,r37.2_net_3,r39_net_1,r39_net_2,r39_net_4,r39_net_5,r39_net_3,r38.1_net,r38.2_net_1,r38.2_net_2,r38.2_net_4,r38.2_net_5,r38.2_net_3,r40_net_1,r40_net_2,r40_net_4,r40_net_5,r40_net_3,r41.1_net_1,r41.1_net_2,r41.1_net_4,r41.1_net_5,r41.1_net_3,r41.2_net_1,r41.2_net_2,r41.2_net_4,r41.2_net_5,r41.2_net_3,r42.1_net_1,r42.1_net_2,r42.1_net_4,r42.1_net_5,r42.1_net_3,r42.2_net_1,r42.2_net_2,r42.2_net_4,r42.2_net_5,r42.2_net_3,r42.4_net_1,r42.4_net_2,r42.4_net_4,r42.4_net_5,r42.4_net_3,r41.3_net,r42.3_net,r63.1_1,r63.1_2,r63.1_3,r63.1_4,r63.1_5,r63.1_6,r63.1_7,r63.1_8,r63.1_9,r63.1_10,r63.1_11,r63.1_12,r63.1_13,r63.2,r64_1,r64_2,r64_3,r64_4,r64_5,r64_6,r64_7,r67.1_1,r67.1_2,r67.1_3,r67.1_4,r67.1_5,r67.1_6,r67.1_7,r68.2a_1,r68.2a_2,r68.2a_3,r68.2a_4,r68.2a_5,r50h1_1,r50h1_2,r50h1_3,r50h1_5,r50h1_6,r50h1_7,r50h1_8,r50h1_12,r50h1_13,r50h1_14,r50h1_15,r50h2_1,r50h2_2,r50h2_3,r50h2_5,r50h2_12,r50h2_13,r50h2_14,r50h2_15,r50h3_1,r50h3_2,r50h3_3,r50h3_5,r50h3_12,r50h3_13,r50h3_14,r50h3_15,r50h4_1,r50h4_2,r50h4_3,r50h4_5,r50h4_12,r50h4_13,r50h4_14,r50h4_15,r50h5_1,r50h5_2,r50h5_3,r50h5_5,r50h5_12,r50h5_13,r50h5_14,r50h5_15";// dashboard variable value

            string[] spss_variable_name = questions.Split(',');
            SpssReader spssDataset;
            using (FileStream fileStream = new FileStream(@"E:\Clue-Box Work\2017\Nov-2017\CANDYCRUSH  TH-MAY-2015\REC\TH-CANDY-2015.sav", FileMode.Open, FileAccess.Read, FileShare.Read, 2048 * 10, FileOptions.SequentialScan))
            {
                spssDataset = new SpssReader(fileStream); // Create the reader, this will read the file header
                foreach (string spss_v in spss_variable_name)
                {
                    foreach (var variable in spssDataset.Variables)  // Iterate through all the varaibles
                    {
                        if (variable.Name.Equals(spss_v))
                        {
                            foreach (KeyValuePair<double, string> label in variable.ValueLabels)
                            {
                                string VARIABLE_NAME = spss_v;
                                string VARIABLE_NAME_SUB_NAME = variable.Name;
                                string VARIABLE_ID = label.Key.ToString();
                                string VARIABLE_VALUE = label.Value;
                                string VARIABLE_NAME_QUESTION = variable.Label;
                                string BASE_VARIABLE_NAME = variable.Name;
                              // iobj.insert_Dashboard_variable_values(VARIABLE_NAME, VARIABLE_NAME_SUB_NAME, VARIABLE_ID, VARIABLE_VALUE, VARIABLE_NAME_QUESTION, SurveyID, Country, BASE_VARIABLE_NAME, SurveyPeriod);
                            }
                        }

                    }
                }
                foreach (var record in spssDataset.Records)
                {
                    string UserID = null;
                    string variable_name;
                    string u_id = null;
                    decimal Weight = 1;
                    string AttendedOn = "-- Not Available --";
                    string Gender = "-- Not Available --";
                    string MaritalStatus = "-- Not Available --";
                    string Location = "-- Not Available --";
                    string AgeGroup = "-- Not Available --";
                    string Education = "-- Not Available --";
                    string Occupation = "-- Not Available --";
                    string PersonalInc_BKK = "-- Not Available --";
                    string HHIncome_BKK = "-- Not Available --";
                    string PersonalInc_UPC = "-- Not Available --";
                    string HHIncome_UPC = "-- Not Available --";
                    string Period = "-- Not Available --";
                    string S5_1 = "-- Not Available --";
                    string S5_2 = "-- Not Available --";
                    string S5_3 = "-- Not Available --";
                    string S5_4 = "-- Not Available --";
                    string S5_5 = "-- Not Available --";
                    string S5_6 = "-- Not Available --";
                    string S5_7 = "-- Not Available --";
                    string S5_8 = "-- Not Available --";
                    string S5_9 = "-- Not Available --";
                    string S5_10 = "-- Not Available --";
                    string S5_11 = "-- Not Available --";
                    string S5_12 = "-- Not Available --";
                    string S5_13 = "-- Not Available --";
                    string S5_14 = "-- Not Available --";
                    string S5_15 = "-- Not Available --";
                    string S5_16 = "-- Not Available --";
                    string S5_17 = "-- Not Available --";
                    string S5_18 = "-- Not Available --";
                    string S5_19 = "-- Not Available --";
                    string S6 = "-- Not Available --";
                    string S7_1 = "-- Not Available --";
                    string S7_2 = "-- Not Available --";
                    string S7_3 = "-- Not Available --";
                    string S7_4 = "-- Not Available --";
                    string S7_5 = "-- Not Available --";
                    string S7_6 = "-- Not Available --";
                    string S7_7 = "-- Not Available --";
                    string S7_8 = "-- Not Available --";
                    string S7_9 = "-- Not Available --";
                    string S7_10 = "-- Not Available --";
                    string S7_11 = "-- Not Available --";
                    string S7_12 = "-- Not Available --";
                    string S7_13 = "-- Not Available --";
                    string S7_14 = "-- Not Available --";
                    string S7_15 = "-- Not Available --";
                    string S7_16 = "-- Not Available --";
                    string S7_17 = "-- Not Available --";
                    string S7_18 = "-- Not Available --";
                    string S7_19 = "-- Not Available --";
                    string S8 = "-- Not Available --";
                    string S10_2 = "-- Not Available --";
                    string S10_3 = "-- Not Available --";
                    string S11_2 = "-- Not Available --";
                    string S11_3 = "-- Not Available --";
                    string BrTom = "-- Not Available --";
                    string BrSpont_Kopipo_org = "-- Not Available --";
                    string BrSpont_Hall_Cool = "-- Not Available --";
                    string BrSpont_Hall_Soother = "-- Not Available --";
                    string BrSpont_Hall_Fruity = "-- Not Available --";
                    string BrSpont_Hall_XS = "-- Not Available --";
                    string BrSpont_Hall_Yellow = "-- Not Available --";
                    string BrSpont_Hall_Green = "-- Not Available --";
                    string BrSpont_Hall_LightBlue = "-- Not Available --";
                    string BrSpont_Hall_HoneyFlav = "-- Not Available --";
                    string BrSpont_Hall_HoneyLemFlav = "-- Not Available --";
                    string BrSpont_Hall_LemonFlavour = "-- Not Available --";
                    string BrSpont_Hall_White = "-- Not Available --";
                    string BrSpont_HartBeat = "-- Not Available --";
                    string BrSpont_KopokoCappuccino = "-- Not Available --";
                    string BrSpont_KopikoSugarFree = "-- Not Available --";
                    string BrSpont_Ole = "-- Not Available --";
                    string BrSpont_Clorets = "-- Not Available --";
                    string BrAid_Kopipo_org = "-- Not Available --";
                    string BrAid_Hall_Cool = "-- Not Available --";
                    string BrAid_Hall_Soother = "-- Not Available --";
                    string BrAid_Hall_Fruity = "-- Not Available --";
                    string BrAid_Hall_XS = "-- Not Available --";
                    string BrAid_Hall_Yellow = "-- Not Available --";
                    string BrAid_Hall_Green = "-- Not Available --";
                    string BrAid_Hall_LightBlue = "-- Not Available --";
                    string BrAid_Hall_HoneyFlav = "-- Not Available --";
                    string BrAid_Hall_HoneyLemFlav = "-- Not Available --";
                    string BrAid_Hall_LemonFlavour = "-- Not Available --";
                    string BrAid_Hall_White = "-- Not Available --";
                    string BrAid_HartBeat = "-- Not Available --";
                    string BrAid_KopokoCappuccino = "-- Not Available --";
                    string BrAid_KopikoSugarFree = "-- Not Available --";
                    string BrAid_Ole = "-- Not Available --";
                    string BrAid_Clorets = "-- Not Available --";
                    string AdTom = "-- Not Available --";
                    string AdSpont_Kopipo_org = "-- Not Available --";
                    string AdSpont_Hall_Cool = "-- Not Available --";
                    string AdSpont_Hall_Soother = "-- Not Available --";
                    string AdSpont_Hall_Fruity = "-- Not Available --";
                    string AdSpont_Hall_XS = "-- Not Available --";
                    string AdSpont_Hall_Yellow = "-- Not Available --";
                    string AdSpont_Hall_Green = "-- Not Available --";
                    string AdSpont_Hall_LightBlue = "-- Not Available --";
                    string AdSpont_Hall_HoneyFlav = "-- Not Available --";
                    string AdSpont_Hall_HoneyLemFlav = "-- Not Available --";
                    string AdSpont_Hall_LemonFlavour = "-- Not Available --";
                    string AdSpont_Hall_White = "-- Not Available --";
                    string AdSpont_HartBeat = "-- Not Available --";
                    string AdSpont_KopokoCappuccino = "-- Not Available --";
                    string AdSpont_KopikoSugarFree = "-- Not Available --";
                    string AdSpont_Ole = "-- Not Available --";
                    string AdSpont_Clorets = "-- Not Available --";
                    string AdAided_Kopipo_org = "-- Not Available --";
                    string AdAided_Hall_Cool = "-- Not Available --";
                    string AdAided_Hall_Soother = "-- Not Available --";
                    string AdAided_Hall_Fruity = "-- Not Available --";
                    string AdAided_Hall_XS = "-- Not Available --";
                    string AdAided_Hall_Yellow = "-- Not Available --";
                    string AdAided_Hall_Green = "-- Not Available --";
                    string AdAided_Hall_LightBlue = "-- Not Available --";
                    string AdAided_Hall_HoneyFlav = "-- Not Available --";
                    string AdAided_Hall_HoneyLemFlav = "-- Not Available --";
                    string AdAided_Hall_LemonFlavour = "-- Not Available --";
                    string AdAided_Hall_White = "-- Not Available --";
                    string AdAided_HartBeat = "-- Not Available --";
                    string AdAided_KopokoCappuccino = "-- Not Available --";
                    string AdAided_KopikoSugarFree = "-- Not Available --";
                    string AdAided_Ole = "-- Not Available --";
                    string AdAided_Clorets = "-- Not Available --";
                    string EverCons_Kopipo_org = "-- Not Available --";
                    string EverCons_Hall_Cool = "-- Not Available --";
                    string EverCons_Hall_Soother = "-- Not Available --";
                    string EverCons_Hall_Fruity = "-- Not Available --";
                    string EverCons_Hall_XS = "-- Not Available --";
                    string EverCons_Hall_Yellow = "-- Not Available --";
                    string EverCons_Hall_Green = "-- Not Available --";
                    string EverCons_Hall_LightBlue = "-- Not Available --";
                    string EverCons_Hall_HoneyFlav = "-- Not Available --";
                    string EverCons_Hall_HoneyLemFlav = "-- Not Available --";
                    string EverCons_Hall_LemonFlavour = "-- Not Available --";
                    string EverCons_Hall_White = "-- Not Available --";
                    string EverCons_HartBeat = "-- Not Available --";
                    string EverCons_KopokoCappuccino = "-- Not Available --";
                    string EverCons_KopikoSugarFree = "-- Not Available --";
                    string EverCons_Ole = "-- Not Available --";
                    string EverCons_Clorets = "-- Not Available --";
                    string ConsL3M_Kopipo_org = "-- Not Available --";
                    string ConsL3M_Hall_Cool = "-- Not Available --";
                    string ConsL3M_Hall_Soother = "-- Not Available --";
                    string ConsL3M_Hall_Fruity = "-- Not Available --";
                    string ConsL3M_Hall_XS = "-- Not Available --";
                    string ConsL3M_Hall_Yellow = "-- Not Available --";
                    string ConsL3M_Hall_Green = "-- Not Available --";
                    string ConsL3M_Hall_LightBlue = "-- Not Available --";
                    string ConsL3M_Hall_HoneyFlav = "-- Not Available --";
                    string ConsL3M_Hall_HoneyLemFlav = "-- Not Available --";
                    string ConsL3M_Hall_LemonFlavour = "-- Not Available --";
                    string ConsL3M_Hall_White = "-- Not Available --";
                    string ConsL3M_HartBeat = "-- Not Available --";
                    string ConsL3M_KopokoCappuccino = "-- Not Available --";
                    string ConsL3M_KopikoSugarFree = "-- Not Available --";
                    string ConsL3M_Ole = "-- Not Available --";
                    string ConsL3M_Clorets = "-- Not Available --";
                    string ConsL1M_Kopipo_org = "-- Not Available --";
                    string ConsL1M_Hall_Cool = "-- Not Available --";
                    string ConsL1M_Hall_Soother = "-- Not Available --";
                    string ConsL1M_Hall_Fruity = "-- Not Available --";
                    string ConsL1M_Hall_XS = "-- Not Available --";
                    string ConsL1M_Hall_Yellow = "-- Not Available --";
                    string ConsL1M_Hall_Green = "-- Not Available --";
                    string ConsL1M_Hall_LightBlue = "-- Not Available --";
                    string ConsL1M_Hall_HoneyFlav = "-- Not Available --";
                    string ConsL1M_Hall_HoneyLemFlav = "-- Not Available --";
                    string ConsL1M_Hall_LemonFlavour = "-- Not Available --";
                    string ConsL1M_Hall_White = "-- Not Available --";
                    string ConsL1M_HartBeat = "-- Not Available --";
                    string ConsL1M_KopokoCappuccino = "-- Not Available --";
                    string ConsL1M_KopikoSugarFree = "-- Not Available --";
                    string ConsL1M_Ole = "-- Not Available --";
                    string ConsL1M_Clorets = "-- Not Available --";
                    string ConsL1W_Kopipo_org = "-- Not Available --";
                    string ConsL1W_Hall_Cool = "-- Not Available --";
                    string ConsL1W_Hall_Soother = "-- Not Available --";
                    string ConsL1W_Hall_Fruity = "-- Not Available --";
                    string ConsL1W_Hall_XS = "-- Not Available --";
                    string ConsL1W_Hall_Yellow = "-- Not Available --";
                    string ConsL1W_Hall_Green = "-- Not Available --";
                    string ConsL1W_Hall_LightBlue = "-- Not Available --";
                    string ConsL1W_Hall_HoneyFlav = "-- Not Available --";
                    string ConsL1W_Hall_HoneyLemFlav = "-- Not Available --";
                    string ConsL1W_Hall_LemonFlavour = "-- Not Available --";
                    string ConsL1W_Hall_White = "-- Not Available --";
                    string ConsL1W_HartBeat = "-- Not Available --";
                    string ConsL1W_KopokoCappuccino = "-- Not Available --";
                    string ConsL1W_KopikoSugarFree = "-- Not Available --";
                    string ConsL1W_Ole = "-- Not Available --";
                    string ConsL1W_Clorets = "-- Not Available --";
                    string IntToBuy_Kopipo_org = "-- Not Available --";
                    string IntToBuy_Hall_Cool = "-- Not Available --";
                    string IntToBuy_Hall_Soother = "-- Not Available --";
                    string IntToBuy_Hall_Fruity = "-- Not Available --";
                    string IntToBuy_Hall_XS = "-- Not Available --";
                    string IntToBuy_Hall_Yellow = "-- Not Available --";
                    string IntToBuy_Hall_Green = "-- Not Available --";
                    string IntToBuy_Hall_LightBlue = "-- Not Available --";
                    string IntToBuy_Hall_HoneyFlav = "-- Not Available --";
                    string IntToBuy_Hall_HoneyLemFlav = "-- Not Available --";
                    string IntToBuy_Hall_LemonFlavour = "-- Not Available --";
                    string IntToBuy_Hall_White = "-- Not Available --";
                    string IntToBuy_HartBeat = "-- Not Available --";
                    string IntToBuy_KopokoCappuccino = "-- Not Available --";
                    string IntToBuy_KopikoSugarFree = "-- Not Available --";
                    string IntToBuy_Ole = "-- Not Available --";
                    string IntToBuy_Clorets = "-- Not Available --";
                    string Bumo = "-- Not Available --";
                    string PreBumo = "-- Not Available --";
                    string NetBrTom = "-- Not Available --";
                    string BrSpontNet_KopikoNet = "-- Not Available --";
                    string BrSpontNet_HallsNet = "-- Not Available --";
                    string BrSpontNet_HartBeatNet = "-- Not Available --";
                    string BrSpontNet_OleNet = "-- Not Available --";
                    string BrSpontNet_CloretsNet = "-- Not Available --";
                    string BrAidNet_KopikoNet = "-- Not Available --";
                    string BrAidNet_HallsNet = "-- Not Available --";
                    string BrAidNet_HartBeatNet = "-- Not Available --";
                    string BrAidNet_OleNet = "-- Not Available --";
                    string BrAidNet_CloretsNet = "-- Not Available --";
                    string NetAdTom = "-- Not Available --";
                    string AdSpontNet_KopikoNet = "-- Not Available --";
                    string AdSpontNet_HallsNet = "-- Not Available --";
                    string AdSpontNet_HartBeatNet = "-- Not Available --";
                    string AdSpontNet_OleNet = "-- Not Available --";
                    string AdSpontNet_CloretsNet = "-- Not Available --";
                    string AdAidNet_KopikoNet = "-- Not Available --";
                    string AdAidNet_HallsNet = "-- Not Available --";
                    string AdAidNet_HartBeatNet = "-- Not Available --";
                    string AdAidNet_OleNet = "-- Not Available --";
                    string AdAidNet_CloretsNet = "-- Not Available --";
                    string EverConsNet_KopikoNet = "-- Not Available --";
                    string EverConsNet_HallsNet = "-- Not Available --";
                    string EverConsNet_HartBeatNet = "-- Not Available --";
                    string EverConsNet_OleNet = "-- Not Available --";
                    string EverConsNet_CloretsNet = "-- Not Available --";
                    string ConsL3MNet_KopikoNet = "-- Not Available --";
                    string ConsL3MNet_HallsNet = "-- Not Available --";
                    string ConsL3MNet_HartBeatNet = "-- Not Available --";
                    string ConsL3MNet_OleNet = "-- Not Available --";
                    string ConsL3MNet_CloretsNet = "-- Not Available --";
                    string ConsL1MNet_KopikoNet = "-- Not Available --";
                    string ConsL1MNet_HallsNet = "-- Not Available --";
                    string ConsL1MNet_HartBeatNet = "-- Not Available --";
                    string ConsL1MNet_OleNet = "-- Not Available --";
                    string ConsL1MNet_CloretsNet = "-- Not Available --";
                    string ConsL1WNet_KopikoNet = "-- Not Available --";
                    string ConsL1WNet_HallsNet = "-- Not Available --";
                    string ConsL1WNet_HartBeatNet = "-- Not Available --";
                    string ConsL1WNet_OleNet = "-- Not Available --";
                    string ConsL1WNet_CloretsNet = "-- Not Available --";
                    string IntBuyNetNet_KopikoNet = "-- Not Available --";
                    string IntBuyNetNet_HallsNet = "-- Not Available --";
                    string IntBuyNetNet_HartBeatNet = "-- Not Available --";
                    string IntBuyNetNet_OleNet = "-- Not Available --";
                    string IntBuyNetNet_CloretsNet = "-- Not Available --";
                    string Bumo_Net = "-- Not Available --";
                    string PreBumo_Net = "-- Not Available --";
                    string r63_1_1 = "-- Not Available --";
                    string r63_1_2 = "-- Not Available --";
                    string r63_1_3 = "-- Not Available --";
                    string r63_1_4 = "-- Not Available --";
                    string r63_1_5 = "-- Not Available --";
                    string r63_1_6 = "-- Not Available --";
                    string r63_1_7 = "-- Not Available --";
                    string r63_1_8 = "-- Not Available --";
                    string r63_1_9 = "-- Not Available --";
                    string r63_1_10 = "-- Not Available --";
                    string r63_1_11 = "-- Not Available --";
                    string r63_1_12 = "-- Not Available --";
                    string r63_1_13 = "-- Not Available --";
                    string r63_2 = "-- Not Available --";
                    string r64_1 = "-- Not Available --";
                    string r64_2 = "-- Not Available --";
                    string r64_3 = "-- Not Available --";
                    string r64_4 = "-- Not Available --";
                    string r64_5 = "-- Not Available --";
                    string r64_6 = "-- Not Available --";
                    string r64_7 = "-- Not Available --";
                    string r67_1_1 = "-- Not Available --";
                    string r67_1_2 = "-- Not Available --";
                    string r67_1_3 = "-- Not Available --";
                    string r67_1_4 = "-- Not Available --";
                    string r67_1_5 = "-- Not Available --";
                    string r67_1_6 = "-- Not Available --";
                    string r67_1_7 = "-- Not Available --";
                    string r68_2a_1 = "-- Not Available --";
                    string r68_2a_2 = "-- Not Available --";
                    string r68_2a_3 = "-- Not Available --";
                    string r68_2a_4 = "-- Not Available --";
                    string r68_2a_5 = "-- Not Available --";
                    string r50h1_1 = "-- Not Available --";
                    string r50h1_2 = "-- Not Available --";
                    string r50h1_3 = "-- Not Available --";
                    string r50h1_5 = "-- Not Available --";
                    string r50h1_6 = "-- Not Available --";
                    string r50h1_7 = "-- Not Available --";
                    string r50h1_8 = "-- Not Available --";
                    string r50h1_12 = "-- Not Available --";
                    string r50h1_13 = "-- Not Available --";
                    string r50h1_14 = "-- Not Available --";
                    string r50h1_15 = "-- Not Available --";
                    string r50h2_1 = "-- Not Available --";
                    string r50h2_2 = "-- Not Available --";
                    string r50h2_3 = "-- Not Available --";
                    string r50h2_5 = "-- Not Available --";
                    string r50h2_12 = "-- Not Available --";
                    string r50h2_13 = "-- Not Available --";
                    string r50h2_14 = "-- Not Available --";
                    string r50h2_15 = "-- Not Available --";
                    string r50h3_1 = "-- Not Available --";
                    string r50h3_2 = "-- Not Available --";
                    string r50h3_3 = "-- Not Available --";
                    string r50h3_5 = "-- Not Available --";
                    string r50h3_12 = "-- Not Available --";
                    string r50h3_13 = "-- Not Available --";
                    string r50h3_14 = "-- Not Available --";
                    string r50h3_15 = "-- Not Available --";
                    string r50h4_1 = "-- Not Available --";
                    string r50h4_2 = "-- Not Available --";
                    string r50h4_3 = "-- Not Available --";
                    string r50h4_5 = "-- Not Available --";
                    string r50h4_12 = "-- Not Available --";
                    string r50h4_13 = "-- Not Available --";
                    string r50h4_14 = "-- Not Available --";
                    string r50h4_15 = "-- Not Available --";
                    string r50h5_1 = "-- Not Available --";
                    string r50h5_2 = "-- Not Available --";
                    string r50h5_3 = "-- Not Available --";
                    string r50h5_5 = "-- Not Available --";
                    string r50h5_12 = "-- Not Available --";
                    string r50h5_13 = "-- Not Available --";
                    string r50h5_14 = "-- Not Available --";
                    string r50h5_15 = "-- Not Available --";



                    foreach (var variable in spssDataset.Variables)
                    {
                        foreach (string spss_v in spss_variable_name)
                        {
                            if (variable.Name.Equals(spss_v))
                            {
                                variable_name = variable.Name;
                                switch (variable_name)
                                {
                                    case "ID":
                                        {
                                            u_id = Convert.ToString(record.GetValue(variable));
                                            UserID = find_UserID(SurveyID, SurveyPeriod, u_id);
                                            //userID = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "Quota_Gender":
                                        {
                                            Gender = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "E3":
                                        {
                                            MaritalStatus = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "Quota_Location":
                                        {
                                            Location = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "Quota_Age":
                                        {
                                            AgeGroup = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "E2":
                                        {
                                            Education = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "E1":
                                        {
                                            Occupation = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S4.1_BKK":
                                        {
                                            PersonalInc_BKK = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S4.2_BKK":
                                        {
                                            HHIncome_BKK = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S4.1_UPC":
                                        {
                                            PersonalInc_UPC = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S4.2_UPC":
                                        {
                                            HHIncome_UPC = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "Wave":
                                        {
                                            Period = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_1":
                                        {
                                            S5_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_2":
                                        {
                                            S5_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_3":
                                        {
                                            S5_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_4":
                                        {
                                            S5_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_5":
                                        {
                                            S5_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_6":
                                        {
                                            S5_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_7":
                                        {
                                            S5_7 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_8":
                                        {
                                            S5_8 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_9":
                                        {
                                            S5_9 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_10":
                                        {
                                            S5_10 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_11":
                                        {
                                            S5_11 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_12":
                                        {
                                            S5_12 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_13":
                                        {
                                            S5_13 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_14":
                                        {
                                            S5_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_15":
                                        {
                                            S5_15 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_16":
                                        {
                                            S5_16 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_17":
                                        {
                                            S5_17 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_18":
                                        {
                                            S5_18 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S5_19":
                                        {
                                            S5_19 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S6":
                                        {
                                            S6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_1":
                                        {
                                            S7_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_2":
                                        {
                                            S7_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_3":
                                        {
                                            S7_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_4":
                                        {
                                            S7_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_5":
                                        {
                                            S7_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_6":
                                        {
                                            S7_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_7":
                                        {
                                            S7_7 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_8":
                                        {
                                            S7_8 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_9":
                                        {
                                            S7_9 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_10":
                                        {
                                            S7_10 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_11":
                                        {
                                            S7_11 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_12":
                                        {
                                            S7_12 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_13":
                                        {
                                            S7_13 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_14":
                                        {
                                            S7_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_15":
                                        {
                                            S7_15 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_16":
                                        {
                                            S7_16 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_17":
                                        {
                                            S7_17 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_18":
                                        {
                                            S7_18 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S7_19":
                                        {
                                            S7_19 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S8":
                                        {
                                            S8 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "S10.2":
                                        {
                                            S10_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S10.3":
                                        {
                                            S10_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S11.2":
                                        {
                                            S11_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "S11.3":
                                        {
                                            S11_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.1":
                                        {
                                            BrTom = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_2":
                                        {
                                            BrSpont_Kopipo_org = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_6":
                                        {
                                            BrSpont_Hall_Cool = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_7":
                                        {
                                            BrSpont_Hall_Soother = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_8":
                                        {
                                            BrSpont_Hall_Fruity = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_9":
                                        {
                                            BrSpont_Hall_XS = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_26":
                                        {
                                            BrSpont_Hall_Yellow = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_27":
                                        {
                                            BrSpont_Hall_Green = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_28":
                                        {
                                            BrSpont_Hall_LightBlue = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_29":
                                        {
                                            BrSpont_Hall_HoneyFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_30":
                                        {
                                            BrSpont_Hall_HoneyLemFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_31":
                                        {
                                            BrSpont_Hall_LemonFlavour = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_35":
                                        {
                                            BrSpont_Hall_White = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_11":
                                        {
                                            BrSpont_HartBeat = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_3":
                                        {
                                            BrSpont_KopokoCappuccino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_4":
                                        {
                                            BrSpont_KopikoSugarFree = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_12":
                                        {
                                            BrSpont_Ole = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_10":
                                        {
                                            BrSpont_Clorets = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_2":
                                        {
                                            BrAid_Kopipo_org = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_6":
                                        {
                                            BrAid_Hall_Cool = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_7":
                                        {
                                            BrAid_Hall_Soother = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_8":
                                        {
                                            BrAid_Hall_Fruity = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_9":
                                        {
                                            BrAid_Hall_XS = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_26":
                                        {
                                            BrAid_Hall_Yellow = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_27":
                                        {
                                            BrAid_Hall_Green = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_28":
                                        {
                                            BrAid_Hall_LightBlue = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_29":
                                        {
                                            BrAid_Hall_HoneyFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_30":
                                        {
                                            BrAid_Hall_HoneyLemFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_31":
                                        {
                                            BrAid_Hall_LemonFlavour = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_35":
                                        {
                                            BrAid_Hall_White = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_11":
                                        {
                                            BrAid_HartBeat = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_3":
                                        {
                                            BrAid_KopokoCappuccino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_4":
                                        {
                                            BrAid_KopikoSugarFree = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_12":
                                        {
                                            BrAid_Ole = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_10":
                                        {
                                            BrAid_Clorets = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.1":
                                        {
                                            AdTom = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_2":
                                        {
                                            AdSpont_Kopipo_org = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_6":
                                        {
                                            AdSpont_Hall_Cool = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_7":
                                        {
                                            AdSpont_Hall_Soother = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_8":
                                        {
                                            AdSpont_Hall_Fruity = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_9":
                                        {
                                            AdSpont_Hall_XS = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_26":
                                        {
                                            AdSpont_Hall_Yellow = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_27":
                                        {
                                            AdSpont_Hall_Green = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_28":
                                        {
                                            AdSpont_Hall_LightBlue = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_29":
                                        {
                                            AdSpont_Hall_HoneyFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_30":
                                        {
                                            AdSpont_Hall_HoneyLemFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_31":
                                        {
                                            AdSpont_Hall_LemonFlavour = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_35":
                                        {
                                            AdSpont_Hall_White = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_11":
                                        {
                                            AdSpont_HartBeat = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_3":
                                        {
                                            AdSpont_KopokoCappuccino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_4":
                                        {
                                            AdSpont_KopikoSugarFree = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_12":
                                        {
                                            AdSpont_Ole = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_10":
                                        {
                                            AdSpont_Clorets = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_2":
                                        {
                                            AdAided_Kopipo_org = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_6":
                                        {
                                            AdAided_Hall_Cool = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_7":
                                        {
                                            AdAided_Hall_Soother = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_8":
                                        {
                                            AdAided_Hall_Fruity = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_9":
                                        {
                                            AdAided_Hall_XS = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_26":
                                        {
                                            AdAided_Hall_Yellow = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_27":
                                        {
                                            AdAided_Hall_Green = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_28":
                                        {
                                            AdAided_Hall_LightBlue = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_29":
                                        {
                                            AdAided_Hall_HoneyFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_30":
                                        {
                                            AdAided_Hall_HoneyLemFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_31":
                                        {
                                            AdAided_Hall_LemonFlavour = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_35":
                                        {
                                            AdAided_Hall_White = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_11":
                                        {
                                            AdAided_HartBeat = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_3":
                                        {
                                            AdAided_KopokoCappuccino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_4":
                                        {
                                            AdAided_KopikoSugarFree = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_12":
                                        {
                                            AdAided_Ole = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_10":
                                        {
                                            AdAided_Clorets = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_2":
                                        {
                                            EverCons_Kopipo_org = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_6":
                                        {
                                            EverCons_Hall_Cool = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_7":
                                        {
                                            EverCons_Hall_Soother = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_8":
                                        {
                                            EverCons_Hall_Fruity = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_9":
                                        {
                                            EverCons_Hall_XS = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_26":
                                        {
                                            EverCons_Hall_Yellow = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_27":
                                        {
                                            EverCons_Hall_Green = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_28":
                                        {
                                            EverCons_Hall_LightBlue = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_29":
                                        {
                                            EverCons_Hall_HoneyFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_30":
                                        {
                                            EverCons_Hall_HoneyLemFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_31":
                                        {
                                            EverCons_Hall_LemonFlavour = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_35":
                                        {
                                            EverCons_Hall_White = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_11":
                                        {
                                            EverCons_HartBeat = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_3":
                                        {
                                            EverCons_KopokoCappuccino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_4":
                                        {
                                            EverCons_KopikoSugarFree = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_12":
                                        {
                                            EverCons_Ole = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_10":
                                        {
                                            EverCons_Clorets = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_2":
                                        {
                                            ConsL3M_Kopipo_org = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_6":
                                        {
                                            ConsL3M_Hall_Cool = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_7":
                                        {
                                            ConsL3M_Hall_Soother = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_8":
                                        {
                                            ConsL3M_Hall_Fruity = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_9":
                                        {
                                            ConsL3M_Hall_XS = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_26":
                                        {
                                            ConsL3M_Hall_Yellow = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_27":
                                        {
                                            ConsL3M_Hall_Green = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_28":
                                        {
                                            ConsL3M_Hall_LightBlue = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_29":
                                        {
                                            ConsL3M_Hall_HoneyFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_30":
                                        {
                                            ConsL3M_Hall_HoneyLemFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_31":
                                        {
                                            ConsL3M_Hall_LemonFlavour = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_35":
                                        {
                                            ConsL3M_Hall_White = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_11":
                                        {
                                            ConsL3M_HartBeat = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_3":
                                        {
                                            ConsL3M_KopokoCappuccino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_4":
                                        {
                                            ConsL3M_KopikoSugarFree = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_12":
                                        {
                                            ConsL3M_Ole = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_10":
                                        {
                                            ConsL3M_Clorets = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_2":
                                        {
                                            ConsL1M_Kopipo_org = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_6":
                                        {
                                            ConsL1M_Hall_Cool = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_7":
                                        {
                                            ConsL1M_Hall_Soother = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_8":
                                        {
                                            ConsL1M_Hall_Fruity = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_9":
                                        {
                                            ConsL1M_Hall_XS = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_26":
                                        {
                                            ConsL1M_Hall_Yellow = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_27":
                                        {
                                            ConsL1M_Hall_Green = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_28":
                                        {
                                            ConsL1M_Hall_LightBlue = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_29":
                                        {
                                            ConsL1M_Hall_HoneyFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_30":
                                        {
                                            ConsL1M_Hall_HoneyLemFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_31":
                                        {
                                            ConsL1M_Hall_LemonFlavour = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_35":
                                        {
                                            ConsL1M_Hall_White = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_11":
                                        {
                                            ConsL1M_HartBeat = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_3":
                                        {
                                            ConsL1M_KopokoCappuccino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_4":
                                        {
                                            ConsL1M_KopikoSugarFree = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_12":
                                        {
                                            ConsL1M_Ole = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_10":
                                        {
                                            ConsL1M_Clorets = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_2":
                                        {
                                            ConsL1W_Kopipo_org = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_6":
                                        {
                                            ConsL1W_Hall_Cool = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_7":
                                        {
                                            ConsL1W_Hall_Soother = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_8":
                                        {
                                            ConsL1W_Hall_Fruity = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_9":
                                        {
                                            ConsL1W_Hall_XS = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_26":
                                        {
                                            ConsL1W_Hall_Yellow = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_27":
                                        {
                                            ConsL1W_Hall_Green = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_28":
                                        {
                                            ConsL1W_Hall_LightBlue = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_29":
                                        {
                                            ConsL1W_Hall_HoneyFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_30":
                                        {
                                            ConsL1W_Hall_HoneyLemFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_31":
                                        {
                                            ConsL1W_Hall_LemonFlavour = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_35":
                                        {
                                            ConsL1W_Hall_White = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_11":
                                        {
                                            ConsL1W_HartBeat = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_3":
                                        {
                                            ConsL1W_KopokoCappuccino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_4":
                                        {
                                            ConsL1W_KopikoSugarFree = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_12":
                                        {
                                            ConsL1W_Ole = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_10":
                                        {
                                            ConsL1W_Clorets = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_2":
                                        {
                                            IntToBuy_Kopipo_org = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_6":
                                        {
                                            IntToBuy_Hall_Cool = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_7":
                                        {
                                            IntToBuy_Hall_Soother = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_8":
                                        {
                                            IntToBuy_Hall_Fruity = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_9":
                                        {
                                            IntToBuy_Hall_XS = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_26":
                                        {
                                            IntToBuy_Hall_Yellow = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_27":
                                        {
                                            IntToBuy_Hall_Green = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_28":
                                        {
                                            IntToBuy_Hall_LightBlue = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_29":
                                        {
                                            IntToBuy_Hall_HoneyFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_30":
                                        {
                                            IntToBuy_Hall_HoneyLemFlav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_31":
                                        {
                                            IntToBuy_Hall_LemonFlavour = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_35":
                                        {
                                            IntToBuy_Hall_White = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_11":
                                        {
                                            IntToBuy_HartBeat = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_3":
                                        {
                                            IntToBuy_KopokoCappuccino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_4":
                                        {
                                            IntToBuy_KopikoSugarFree = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_12":
                                        {
                                            IntToBuy_Ole = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_10":
                                        {
                                            IntToBuy_Clorets = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.3":
                                        {
                                            Bumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.3":
                                        {
                                            PreBumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.1_net":
                                        {
                                            NetBrTom = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_net_1":
                                        {
                                            BrSpontNet_KopikoNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_net_2":
                                        {
                                            BrSpontNet_HallsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_net_4":
                                        {
                                            BrSpontNet_HartBeatNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_net_5":
                                        {
                                            BrSpontNet_OleNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37.2_net_3":
                                        {
                                            BrSpontNet_CloretsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_net_1":
                                        {
                                            BrAidNet_KopikoNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_net_2":
                                        {
                                            BrAidNet_HallsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_net_4":
                                        {
                                            BrAidNet_HartBeatNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_net_5":
                                        {
                                            BrAidNet_OleNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r39_net_3":
                                        {
                                            BrAidNet_CloretsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.1_net":
                                        {
                                            NetAdTom = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_net_1":
                                        {
                                            AdSpontNet_KopikoNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_net_2":
                                        {
                                            AdSpontNet_HallsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_net_4":
                                        {
                                            AdSpontNet_HartBeatNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_net_5":
                                        {
                                            AdSpontNet_OleNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r38.2_net_3":
                                        {
                                            AdSpontNet_CloretsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_net_1":
                                        {
                                            AdAidNet_KopikoNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_net_2":
                                        {
                                            AdAidNet_HallsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_net_4":
                                        {
                                            AdAidNet_HartBeatNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_net_5":
                                        {
                                            AdAidNet_OleNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r40_net_3":
                                        {
                                            AdAidNet_CloretsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_net_1":
                                        {
                                            EverConsNet_KopikoNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_net_2":
                                        {
                                            EverConsNet_HallsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_net_4":
                                        {
                                            EverConsNet_HartBeatNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_net_5":
                                        {
                                            EverConsNet_OleNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.1_net_3":
                                        {
                                            EverConsNet_CloretsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_net_1":
                                        {
                                            ConsL3MNet_KopikoNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_net_2":
                                        {
                                            ConsL3MNet_HallsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_net_4":
                                        {
                                            ConsL3MNet_HartBeatNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_net_5":
                                        {
                                            ConsL3MNet_OleNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.2_net_3":
                                        {
                                            ConsL3MNet_CloretsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_net_1":
                                        {
                                            ConsL1MNet_KopikoNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_net_2":
                                        {
                                            ConsL1MNet_HallsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_net_4":
                                        {
                                            ConsL1MNet_HartBeatNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_net_5":
                                        {
                                            ConsL1MNet_OleNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.1_net_3":
                                        {
                                            ConsL1MNet_CloretsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_net_1":
                                        {
                                            ConsL1WNet_KopikoNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_net_2":
                                        {
                                            ConsL1WNet_HallsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_net_4":
                                        {
                                            ConsL1WNet_HartBeatNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_net_5":
                                        {
                                            ConsL1WNet_OleNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.2_net_3":
                                        {
                                            ConsL1WNet_CloretsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_net_1":
                                        {
                                            IntBuyNetNet_KopikoNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_net_2":
                                        {
                                            IntBuyNetNet_HallsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_net_4":
                                        {
                                            IntBuyNetNet_HartBeatNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_net_5":
                                        {
                                            IntBuyNetNet_OleNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.4_net_3":
                                        {
                                            IntBuyNetNet_CloretsNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r41.3_net":
                                        {
                                            Bumo_Net = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r42.3_net":
                                        {
                                            PreBumo_Net = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r63.1_1":
                                        {
                                            r63_1_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r63.1_2":
                                        {
                                            r63_1_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r63.1_3":
                                        {
                                            r63_1_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r63.1_4":
                                        {
                                            r63_1_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r63.1_5":
                                        {
                                            r63_1_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r63.1_6":
                                        {
                                            r63_1_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r63.1_7":
                                        {
                                            r63_1_7 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r63.1_8":
                                        {
                                            r63_1_8 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r63.1_9":
                                        {
                                            r63_1_9 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r63.1_10":
                                        {
                                            r63_1_10 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r63.1_11":
                                        {
                                            r63_1_11 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r63.1_12":
                                        {
                                            r63_1_12 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r63.1_13":
                                        {
                                            r63_1_13 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r63.2":
                                        {
                                            r63_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r64_1":
                                        {
                                            r64_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r64_2":
                                        {
                                            r64_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r64_3":
                                        {
                                            r64_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r64_4":
                                        {
                                            r64_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r64_5":
                                        {
                                            r64_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r64_6":
                                        {
                                            r64_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r64_7":
                                        {
                                            r64_7 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r67.1_1":
                                        {
                                            r67_1_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r67.1_2":
                                        {
                                            r67_1_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r67.1_3":
                                        {
                                            r67_1_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r67.1_4":
                                        {
                                            r67_1_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r67.1_5":
                                        {
                                            r67_1_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r67.1_6":
                                        {
                                            r67_1_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r67.1_7":
                                        {
                                            r67_1_7 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r68.2a_1":
                                        {
                                            r68_2a_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r68.2a_2":
                                        {
                                            r68_2a_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r68.2a_3":
                                        {
                                            r68_2a_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r68.2a_4":
                                        {
                                            r68_2a_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r68.2a_5":
                                        {
                                            r68_2a_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h1_1":
                                        {
                                            r50h1_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h1_2":
                                        {
                                            r50h1_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h1_3":
                                        {
                                            r50h1_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h1_5":
                                        {
                                            r50h1_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h1_6":
                                        {
                                            r50h1_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h1_7":
                                        {
                                            r50h1_7 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h1_8":
                                        {
                                            r50h1_8 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h1_12":
                                        {
                                            r50h1_12 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h1_13":
                                        {
                                            r50h1_13 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h1_14":
                                        {
                                            r50h1_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h1_15":
                                        {
                                            r50h1_15 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h2_1":
                                        {
                                            r50h2_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h2_2":
                                        {
                                            r50h2_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h2_3":
                                        {
                                            r50h2_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h2_5":
                                        {
                                            r50h2_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h2_12":
                                        {
                                            r50h2_12 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h2_13":
                                        {
                                            r50h2_13 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h2_14":
                                        {
                                            r50h2_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h2_15":
                                        {
                                            r50h2_15 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h3_1":
                                        {
                                            r50h3_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h3_2":
                                        {
                                            r50h3_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h3_3":
                                        {
                                            r50h3_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h3_5":
                                        {
                                            r50h3_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h3_12":
                                        {
                                            r50h3_12 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h3_13":
                                        {
                                            r50h3_13 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h3_14":
                                        {
                                            r50h3_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h3_15":
                                        {
                                            r50h3_15 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h4_1":
                                        {
                                            r50h4_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h4_2":
                                        {
                                            r50h4_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h4_3":
                                        {
                                            r50h4_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h4_5":
                                        {
                                            r50h4_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h4_12":
                                        {
                                            r50h4_12 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h4_13":
                                        {
                                            r50h4_13 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h4_14":
                                        {
                                            r50h4_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h4_15":
                                        {
                                            r50h4_15 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h5_1":
                                        {
                                            r50h5_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h5_2":
                                        {
                                            r50h5_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h5_3":
                                        {
                                            r50h5_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h5_5":
                                        {
                                            r50h5_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h5_12":
                                        {
                                            r50h5_12 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h5_13":
                                        {
                                            r50h5_13 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h5_14":
                                        {
                                            r50h5_14 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r50h5_15":
                                        {
                                            r50h5_15 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }


                                }
                            }
                        }
                    }
                    if (UserID != null && Weight != 0)
                    {
                        iobj.insert_Dashboard_values(UserID, Country, SurveyID, Gender, MaritalStatus, SurveyPeriod, Weight, Location, AgeGroup, Education, Occupation, PersonalInc_BKK, HHIncome_BKK, PersonalInc_UPC, HHIncome_UPC, Period, S5_1, S5_2, S5_3, S5_4, S5_5, S5_6, S5_7, S5_8, S5_9, S5_10, S5_11, S5_12, S5_13, S5_14, S5_15, S5_16, S5_17, S5_18, S5_19, S6, S7_1, S7_2, S7_3, S7_4, S7_5, S7_6, S7_7, S7_8, S7_9, S7_10, S7_11, S7_12, S7_13, S7_14, S7_15, S7_16, S7_17, S7_18, S7_19, S8, S10_2, S10_3, S11_2, S11_3, BrTom, BrSpont_Kopipo_org, BrSpont_Hall_Cool, BrSpont_Hall_Soother, BrSpont_Hall_Fruity, BrSpont_Hall_XS, BrSpont_Hall_Yellow, BrSpont_Hall_Green, BrSpont_Hall_LightBlue, BrSpont_Hall_HoneyFlav, BrSpont_Hall_HoneyLemFlav, BrSpont_Hall_LemonFlavour, BrSpont_Hall_White, BrSpont_HartBeat, BrSpont_KopokoCappuccino, BrSpont_KopikoSugarFree, BrSpont_Ole, BrSpont_Clorets, BrAid_Kopipo_org, BrAid_Hall_Cool, BrAid_Hall_Soother, BrAid_Hall_Fruity, BrAid_Hall_XS, BrAid_Hall_Yellow, BrAid_Hall_Green, BrAid_Hall_LightBlue, BrAid_Hall_HoneyFlav, BrAid_Hall_HoneyLemFlav, BrAid_Hall_LemonFlavour, BrAid_Hall_White, BrAid_HartBeat, BrAid_KopokoCappuccino, BrAid_KopikoSugarFree, BrAid_Ole, BrAid_Clorets, AdTom, AdSpont_Kopipo_org, AdSpont_Hall_Cool, AdSpont_Hall_Soother, AdSpont_Hall_Fruity, AdSpont_Hall_XS, AdSpont_Hall_Yellow, AdSpont_Hall_Green, AdSpont_Hall_LightBlue, AdSpont_Hall_HoneyFlav, AdSpont_Hall_HoneyLemFlav, AdSpont_Hall_LemonFlavour, AdSpont_Hall_White, AdSpont_HartBeat, AdSpont_KopokoCappuccino, AdSpont_KopikoSugarFree, AdSpont_Ole, AdSpont_Clorets, AdAided_Kopipo_org, AdAided_Hall_Cool, AdAided_Hall_Soother, AdAided_Hall_Fruity, AdAided_Hall_XS, AdAided_Hall_Yellow, AdAided_Hall_Green, AdAided_Hall_LightBlue, AdAided_Hall_HoneyFlav, AdAided_Hall_HoneyLemFlav, AdAided_Hall_LemonFlavour, AdAided_Hall_White, AdAided_HartBeat, AdAided_KopokoCappuccino, AdAided_KopikoSugarFree, AdAided_Ole, AdAided_Clorets, EverCons_Kopipo_org, EverCons_Hall_Cool, EverCons_Hall_Soother, EverCons_Hall_Fruity, EverCons_Hall_XS, EverCons_Hall_Yellow, EverCons_Hall_Green, EverCons_Hall_LightBlue, EverCons_Hall_HoneyFlav, EverCons_Hall_HoneyLemFlav, EverCons_Hall_LemonFlavour, EverCons_Hall_White, EverCons_HartBeat, EverCons_KopokoCappuccino, EverCons_KopikoSugarFree, EverCons_Ole, EverCons_Clorets, ConsL3M_Kopipo_org, ConsL3M_Hall_Cool, ConsL3M_Hall_Soother, ConsL3M_Hall_Fruity, ConsL3M_Hall_XS, ConsL3M_Hall_Yellow, ConsL3M_Hall_Green, ConsL3M_Hall_LightBlue, ConsL3M_Hall_HoneyFlav, ConsL3M_Hall_HoneyLemFlav, ConsL3M_Hall_LemonFlavour, ConsL3M_Hall_White, ConsL3M_HartBeat, ConsL3M_KopokoCappuccino, ConsL3M_KopikoSugarFree, ConsL3M_Ole, ConsL3M_Clorets, ConsL1M_Kopipo_org, ConsL1M_Hall_Cool, ConsL1M_Hall_Soother, ConsL1M_Hall_Fruity, ConsL1M_Hall_XS, ConsL1M_Hall_Yellow, ConsL1M_Hall_Green, ConsL1M_Hall_LightBlue, ConsL1M_Hall_HoneyFlav, ConsL1M_Hall_HoneyLemFlav, ConsL1M_Hall_LemonFlavour, ConsL1M_Hall_White, ConsL1M_HartBeat, ConsL1M_KopokoCappuccino, ConsL1M_KopikoSugarFree, ConsL1M_Ole, ConsL1M_Clorets, ConsL1W_Kopipo_org, ConsL1W_Hall_Cool, ConsL1W_Hall_Soother, ConsL1W_Hall_Fruity, ConsL1W_Hall_XS, ConsL1W_Hall_Yellow, ConsL1W_Hall_Green, ConsL1W_Hall_LightBlue, ConsL1W_Hall_HoneyFlav, ConsL1W_Hall_HoneyLemFlav, ConsL1W_Hall_LemonFlavour, ConsL1W_Hall_White, ConsL1W_HartBeat, ConsL1W_KopokoCappuccino, ConsL1W_KopikoSugarFree, ConsL1W_Ole, ConsL1W_Clorets, IntToBuy_Kopipo_org, IntToBuy_Hall_Cool, IntToBuy_Hall_Soother, IntToBuy_Hall_Fruity, IntToBuy_Hall_XS, IntToBuy_Hall_Yellow, IntToBuy_Hall_Green, IntToBuy_Hall_LightBlue, IntToBuy_Hall_HoneyFlav, IntToBuy_Hall_HoneyLemFlav, IntToBuy_Hall_LemonFlavour, IntToBuy_Hall_White, IntToBuy_HartBeat, IntToBuy_KopokoCappuccino, IntToBuy_KopikoSugarFree, IntToBuy_Ole, IntToBuy_Clorets, Bumo, PreBumo, NetBrTom, BrSpontNet_KopikoNet, BrSpontNet_HallsNet, BrSpontNet_HartBeatNet, BrSpontNet_OleNet, BrSpontNet_CloretsNet, BrAidNet_KopikoNet, BrAidNet_HallsNet, BrAidNet_HartBeatNet, BrAidNet_OleNet, BrAidNet_CloretsNet, NetAdTom, AdSpontNet_KopikoNet, AdSpontNet_HallsNet, AdSpontNet_HartBeatNet, AdSpontNet_OleNet, AdSpontNet_CloretsNet, AdAidNet_KopikoNet, AdAidNet_HallsNet, AdAidNet_HartBeatNet, AdAidNet_OleNet, AdAidNet_CloretsNet, EverConsNet_KopikoNet, EverConsNet_HallsNet, EverConsNet_HartBeatNet, EverConsNet_OleNet, EverConsNet_CloretsNet, ConsL3MNet_KopikoNet, ConsL3MNet_HallsNet, ConsL3MNet_HartBeatNet, ConsL3MNet_OleNet, ConsL3MNet_CloretsNet, ConsL1MNet_KopikoNet, ConsL1MNet_HallsNet, ConsL1MNet_HartBeatNet, ConsL1MNet_OleNet, ConsL1MNet_CloretsNet, ConsL1WNet_KopikoNet, ConsL1WNet_HallsNet, ConsL1WNet_HartBeatNet, ConsL1WNet_OleNet, ConsL1WNet_CloretsNet, IntBuyNetNet_KopikoNet, IntBuyNetNet_HallsNet, IntBuyNetNet_HartBeatNet, IntBuyNetNet_OleNet, IntBuyNetNet_CloretsNet, Bumo_Net, PreBumo_Net, r63_1_1, r63_1_2, r63_1_3, r63_1_4, r63_1_5, r63_1_6, r63_1_7, r63_1_8, r63_1_9, r63_1_10, r63_1_11, r63_1_12, r63_1_13, r63_2, r64_1, r64_2, r64_3, r64_4, r64_5, r64_6, r64_7, r67_1_1, r67_1_2, r67_1_3, r67_1_4, r67_1_5, r67_1_6, r67_1_7, r68_2a_1, r68_2a_2, r68_2a_3, r68_2a_4, r68_2a_5, r50h1_1, r50h1_2, r50h1_3, r50h1_5, r50h1_6, r50h1_7, r50h1_8, r50h1_12, r50h1_13, r50h1_14, r50h1_15, r50h2_1, r50h2_2, r50h2_3, r50h2_5, r50h2_12, r50h2_13, r50h2_14, r50h2_15, r50h3_1, r50h3_2, r50h3_3, r50h3_5, r50h3_12, r50h3_13, r50h3_14, r50h3_15, r50h4_1, r50h4_2, r50h4_3, r50h4_5, r50h4_12, r50h4_13, r50h4_14, r50h4_15, r50h5_1, r50h5_2, r50h5_3, r50h5_5, r50h5_12, r50h5_13, r50h5_14, r50h5_15);
                    }

                }

            }
        }

        private static string getYear(string SurveyPeriod)
        {
            string[] date = SurveyPeriod.Split('-');
            return date[0];
        }

        private static string find_UserID(int SurveyID, string SurveyPeriod, string u_id)
        {
            string sum = "";
            string[] date = SurveyPeriod.Split('-');
            foreach (string d in date)
            {
                sum = sum + d;

            }
            return u_id + SurveyID + sum;
        }

    }


    class ThInsertion
    {
        SqlConnection connection = new SqlConnection("Data Source=52.74.59.117;Initial Catalog=ClueboxMobile;Persist Security Info=True;User ID=sa;Password=ClueBox123*;");
        internal void insert_Dashboard_variable_values(string VARIABLE_NAME, string VARIABLE_NAME_SUB_NAME, string VARIABLE_ID, string VARIABLE_VALUE, string VARIABLE_NAME_QUESTION, int SurveyID, string country, string BASE_VARIABLE_NAME, string SurveyPeriod)
        {
            connection.Open();
            SqlCommand cmd = new SqlCommand("INSERT INTO DashboardSPS_Variable_Values (VARIABLE_NAME,VARIABLE_NAME_SUB_NAME,VARIABLE_ID,VARIABLE_VALUE,VARIABLE_NAME_QUESTION,SURVEY_ID,SURVEY_COUNTRY,BASE_VARIABLE_NAME,SURVEY_PERIOD) " + "VALUES('" + VARIABLE_NAME + "','" + VARIABLE_NAME_SUB_NAME + "','" + VARIABLE_ID + "','" + VARIABLE_VALUE.Replace("'", "''") + "','" + VARIABLE_NAME_QUESTION + "','" + SurveyID + "','" + country + "','" + BASE_VARIABLE_NAME + "','" + SurveyPeriod + "')", connection);
            try
            {


                cmd.ExecuteNonQuery();
                Console.WriteLine("Dashboard variable values inserted successfully");

                connection.Close();



            }
            catch (Exception)
            {

                Console.WriteLine("Exception occured");
                Console.ReadLine();
            }

        }


        internal void insert_Dashboard_values(string UserID, string Country, int SurveyID, string Gender, string MaritalStatus, string SurveyPeriod, decimal Weight, string Location, string AgeGroup, string Education, string Occupation, string PersonalInc_BKK, string HHIncome_BKK, string PersonalInc_UPC, string HHIncome_UPC, string Period, string S5_1, string S5_2, string S5_3, string S5_4, string S5_5, string S5_6, string S5_7, string S5_8, string S5_9, string S5_10, string S5_11, string S5_12, string S5_13, string S5_14, string S5_15, string S5_16, string S5_17, string S5_18, string S5_19, string S6, string S7_1, string S7_2, string S7_3, string S7_4, string S7_5, string S7_6, string S7_7, string S7_8, string S7_9, string S7_10, string S7_11, string S7_12, string S7_13, string S7_14, string S7_15, string S7_16, string S7_17, string S7_18, string S7_19, string S8, string S10_2, string S10_3, string S11_2, string S11_3, string BrTom, string BrSpont_Kopipo_org, string BrSpont_Hall_Cool, string BrSpont_Hall_Soother, string BrSpont_Hall_Fruity, string BrSpont_Hall_XS, string BrSpont_Hall_Yellow, string BrSpont_Hall_Green, string BrSpont_Hall_LightBlue, string BrSpont_Hall_HoneyFlav, string BrSpont_Hall_HoneyLemFlav, string BrSpont_Hall_LemonFlavour, string BrSpont_Hall_White, string BrSpont_HartBeat, string BrSpont_KopokoCappuccino, string BrSpont_KopikoSugarFree, string BrSpont_Ole, string BrSpont_Clorets, string BrAid_Kopipo_org, string BrAid_Hall_Cool, string BrAid_Hall_Soother, string BrAid_Hall_Fruity, string BrAid_Hall_XS, string BrAid_Hall_Yellow, string BrAid_Hall_Green, string BrAid_Hall_LightBlue, string BrAid_Hall_HoneyFlav, string BrAid_Hall_HoneyLemFlav, string BrAid_Hall_LemonFlavour, string BrAid_Hall_White, string BrAid_HartBeat, string BrAid_KopokoCappuccino, string BrAid_KopikoSugarFree, string BrAid_Ole, string BrAid_Clorets, string AdTom, string AdSpont_Kopipo_org, string AdSpont_Hall_Cool, string AdSpont_Hall_Soother, string AdSpont_Hall_Fruity, string AdSpont_Hall_XS, string AdSpont_Hall_Yellow, string AdSpont_Hall_Green, string AdSpont_Hall_LightBlue, string AdSpont_Hall_HoneyFlav, string AdSpont_Hall_HoneyLemFlav, string AdSpont_Hall_LemonFlavour, string AdSpont_Hall_White, string AdSpont_HartBeat, string AdSpont_KopokoCappuccino, string AdSpont_KopikoSugarFree, string AdSpont_Ole, string AdSpont_Clorets, string AdAided_Kopipo_org, string AdAided_Hall_Cool, string AdAided_Hall_Soother, string AdAided_Hall_Fruity, string AdAided_Hall_XS, string AdAided_Hall_Yellow, string AdAided_Hall_Green, string AdAided_Hall_LightBlue, string AdAided_Hall_HoneyFlav, string AdAided_Hall_HoneyLemFlav, string AdAided_Hall_LemonFlavour, string AdAided_Hall_White, string AdAided_HartBeat, string AdAided_KopokoCappuccino, string AdAided_KopikoSugarFree, string AdAided_Ole, string AdAided_Clorets, string EverCons_Kopipo_org, string EverCons_Hall_Cool, string EverCons_Hall_Soother, string EverCons_Hall_Fruity, string EverCons_Hall_XS, string EverCons_Hall_Yellow, string EverCons_Hall_Green, string EverCons_Hall_LightBlue, string EverCons_Hall_HoneyFlav, string EverCons_Hall_HoneyLemFlav, string EverCons_Hall_LemonFlavour, string EverCons_Hall_White, string EverCons_HartBeat, string EverCons_KopokoCappuccino, string EverCons_KopikoSugarFree, string EverCons_Ole, string EverCons_Clorets, string ConsL3M_Kopipo_org, string ConsL3M_Hall_Cool, string ConsL3M_Hall_Soother, string ConsL3M_Hall_Fruity, string ConsL3M_Hall_XS, string ConsL3M_Hall_Yellow, string ConsL3M_Hall_Green, string ConsL3M_Hall_LightBlue, string ConsL3M_Hall_HoneyFlav, string ConsL3M_Hall_HoneyLemFlav, string ConsL3M_Hall_LemonFlavour, string ConsL3M_Hall_White, string ConsL3M_HartBeat, string ConsL3M_KopokoCappuccino, string ConsL3M_KopikoSugarFree, string ConsL3M_Ole, string ConsL3M_Clorets, string ConsL1M_Kopipo_org, string ConsL1M_Hall_Cool, string ConsL1M_Hall_Soother, string ConsL1M_Hall_Fruity, string ConsL1M_Hall_XS, string ConsL1M_Hall_Yellow, string ConsL1M_Hall_Green, string ConsL1M_Hall_LightBlue, string ConsL1M_Hall_HoneyFlav, string ConsL1M_Hall_HoneyLemFlav, string ConsL1M_Hall_LemonFlavour, string ConsL1M_Hall_White, string ConsL1M_HartBeat, string ConsL1M_KopokoCappuccino, string ConsL1M_KopikoSugarFree, string ConsL1M_Ole, string ConsL1M_Clorets, string ConsL1W_Kopipo_org, string ConsL1W_Hall_Cool, string ConsL1W_Hall_Soother, string ConsL1W_Hall_Fruity, string ConsL1W_Hall_XS, string ConsL1W_Hall_Yellow, string ConsL1W_Hall_Green, string ConsL1W_Hall_LightBlue, string ConsL1W_Hall_HoneyFlav, string ConsL1W_Hall_HoneyLemFlav, string ConsL1W_Hall_LemonFlavour, string ConsL1W_Hall_White, string ConsL1W_HartBeat, string ConsL1W_KopokoCappuccino, string ConsL1W_KopikoSugarFree, string ConsL1W_Ole, string ConsL1W_Clorets, string IntToBuy_Kopipo_org, string IntToBuy_Hall_Cool, string IntToBuy_Hall_Soother, string IntToBuy_Hall_Fruity, string IntToBuy_Hall_XS, string IntToBuy_Hall_Yellow, string IntToBuy_Hall_Green, string IntToBuy_Hall_LightBlue, string IntToBuy_Hall_HoneyFlav, string IntToBuy_Hall_HoneyLemFlav, string IntToBuy_Hall_LemonFlavour, string IntToBuy_Hall_White, string IntToBuy_HartBeat, string IntToBuy_KopokoCappuccino, string IntToBuy_KopikoSugarFree, string IntToBuy_Ole, string IntToBuy_Clorets, string Bumo, string PreBumo, string NetBrTom, string BrSpontNet_KopikoNet, string BrSpontNet_HallsNet, string BrSpontNet_HartBeatNet, string BrSpontNet_OleNet, string BrSpontNet_CloretsNet, string BrAidNet_KopikoNet, string BrAidNet_HallsNet, string BrAidNet_HartBeatNet, string BrAidNet_OleNet, string BrAidNet_CloretsNet, string NetAdTom, string AdSpontNet_KopikoNet, string AdSpontNet_HallsNet, string AdSpontNet_HartBeatNet, string AdSpontNet_OleNet, string AdSpontNet_CloretsNet, string AdAidNet_KopikoNet, string AdAidNet_HallsNet, string AdAidNet_HartBeatNet, string AdAidNet_OleNet, string AdAidNet_CloretsNet, string EverConsNet_KopikoNet, string EverConsNet_HallsNet, string EverConsNet_HartBeatNet, string EverConsNet_OleNet, string EverConsNet_CloretsNet, string ConsL3MNet_KopikoNet, string ConsL3MNet_HallsNet, string ConsL3MNet_HartBeatNet, string ConsL3MNet_OleNet, string ConsL3MNet_CloretsNet, string ConsL1MNet_KopikoNet, string ConsL1MNet_HallsNet, string ConsL1MNet_HartBeatNet, string ConsL1MNet_OleNet, string ConsL1MNet_CloretsNet, string ConsL1WNet_KopikoNet, string ConsL1WNet_HallsNet, string ConsL1WNet_HartBeatNet, string ConsL1WNet_OleNet, string ConsL1WNet_CloretsNet, string IntBuyNetNet_KopikoNet, string IntBuyNetNet_HallsNet, string IntBuyNetNet_HartBeatNet, string IntBuyNetNet_OleNet, string IntBuyNetNet_CloretsNet, string Bumo_Net, string PreBumo_Net, string r63_1_1, string r63_1_2, string r63_1_3, string r63_1_4, string r63_1_5, string r63_1_6, string r63_1_7, string r63_1_8, string r63_1_9, string r63_1_10, string r63_1_11, string r63_1_12, string r63_1_13, string r63_2, string r64_1, string r64_2, string r64_3, string r64_4, string r64_5, string r64_6, string r64_7, string r67_1_1, string r67_1_2, string r67_1_3, string r67_1_4, string r67_1_5, string r67_1_6, string r67_1_7, string r68_2a_1, string r68_2a_2, string r68_2a_3, string r68_2a_4, string r68_2a_5, string r50h1_1, string r50h1_2, string r50h1_3, string r50h1_5, string r50h1_6, string r50h1_7, string r50h1_8, string r50h1_12, string r50h1_13, string r50h1_14, string r50h1_15, string r50h2_1, string r50h2_2, string r50h2_3, string r50h2_5, string r50h2_12, string r50h2_13, string r50h2_14, string r50h2_15, string r50h3_1, string r50h3_2, string r50h3_3, string r50h3_5, string r50h3_12, string r50h3_13, string r50h3_14, string r50h3_15, string r50h4_1, string r50h4_2, string r50h4_3, string r50h4_5, string r50h4_12, string r50h4_13, string r50h4_14, string r50h4_15, string r50h5_1, string r50h5_2, string r50h5_3, string r50h5_5, string r50h5_12, string r50h5_13, string r50h5_14, string r50h5_15)
        {
            int i;
            connection.Open();
            //SqlConnection connection = new SqlConnection("Data Source=52.74.59.117;Initial Catalog=ClueboxMobile;Persist Security Info=True;User ID=sa;Password=ClueBox123*;");
            SqlCommand cmd = new SqlCommand("INSERT INTO DashboardFlat_THCANDY_temp (UserID,Country,SurveyID,Gender,MaritalStatus,AttendedOn,Weight,Location,AgeGroup,Education,Occupation,PersonalInc_BKK,HHIncome_BKK,PersonalInc_UPC,HHIncome_UPC,Period,S5_1,S5_2,S5_3,S5_4,S5_5,S5_6,S5_7,S5_8,S5_9,S5_10,S5_11,S5_12,S5_13,S5_14,S5_15,S5_16,S5_17,S5_18,S5_19,S6,S7_1,S7_2,S7_3,S7_4,S7_5,S7_6,S7_7,S7_8,S7_9,S7_10,S7_11,S7_12,S7_13,S7_14,S7_15,S7_16,S7_17,S7_18,S7_19,S8,S10_2,S10_3,S11_2,S11_3,BrTom,BrSpont_Kopipo_org,BrSpont_Hall_Cool,BrSpont_Hall_Soother,BrSpont_Hall_Fruity,BrSpont_Hall_XS,BrSpont_Hall_Yellow,BrSpont_Hall_Green,BrSpont_Hall_LightBlue,BrSpont_Hall_HoneyFlav,BrSpont_Hall_HoneyLemFlav,BrSpont_Hall_LemonFlavour,BrSpont_Hall_White,BrSpont_HartBeat,BrSpont_KopokoCappuccino,BrSpont_KopikoSugarFree,BrSpont_Ole,BrSpont_Clorets,BrAid_Kopipo_org,BrAid_Hall_Cool,BrAid_Hall_Soother,BrAid_Hall_Fruity,BrAid_Hall_XS,BrAid_Hall_Yellow,BrAid_Hall_Green,BrAid_Hall_LightBlue,BrAid_Hall_HoneyFlav,BrAid_Hall_HoneyLemFlav,BrAid_Hall_LemonFlavour,BrAid_Hall_White,BrAid_HartBeat,BrAid_KopokoCappuccino,BrAid_KopikoSugarFree,BrAid_Ole,BrAid_Clorets,AdTom,AdSpont_Kopipo_org,AdSpont_Hall_Cool,AdSpont_Hall_Soother,AdSpont_Hall_Fruity,AdSpont_Hall_XS,AdSpont_Hall_Yellow,AdSpont_Hall_Green,AdSpont_Hall_LightBlue,AdSpont_Hall_HoneyFlav,AdSpont_Hall_HoneyLemFlav,AdSpont_Hall_LemonFlavour,AdSpont_Hall_White,AdSpont_HartBeat,AdSpont_KopokoCappuccino,AdSpont_KopikoSugarFree,AdSpont_Ole,AdSpont_Clorets,AdAided_Kopipo_org,AdAided_Hall_Cool,AdAided_Hall_Soother,AdAided_Hall_Fruity,AdAided_Hall_XS,AdAided_Hall_Yellow,AdAided_Hall_Green,AdAided_Hall_LightBlue,AdAided_Hall_HoneyFlav,AdAided_Hall_HoneyLemFlav,AdAided_Hall_LemonFlavour,AdAided_Hall_White,AdAided_HartBeat,AdAided_KopokoCappuccino,AdAided_KopikoSugarFree,AdAided_Ole,AdAided_Clorets,EverCons_Kopipo_org,EverCons_Hall_Cool,EverCons_Hall_Soother,EverCons_Hall_Fruity,EverCons_Hall_XS,EverCons_Hall_Yellow,EverCons_Hall_Green,EverCons_Hall_LightBlue,EverCons_Hall_HoneyFlav,EverCons_Hall_HoneyLemFlav,EverCons_Hall_LemonFlavour,EverCons_Hall_White,EverCons_HartBeat,EverCons_KopokoCappuccino,EverCons_KopikoSugarFree,EverCons_Ole,EverCons_Clorets,ConsL3M_Kopipo_org,ConsL3M_Hall_Cool,ConsL3M_Hall_Soother,ConsL3M_Hall_Fruity,ConsL3M_Hall_XS,ConsL3M_Hall_Yellow,ConsL3M_Hall_Green,ConsL3M_Hall_LightBlue,ConsL3M_Hall_HoneyFlav,ConsL3M_Hall_HoneyLemFlav,ConsL3M_Hall_LemonFlavour,ConsL3M_Hall_White,ConsL3M_HartBeat,ConsL3M_KopokoCappuccino,ConsL3M_KopikoSugarFree,ConsL3M_Ole,ConsL3M_Clorets,ConsL1M_Kopipo_org,ConsL1M_Hall_Cool,ConsL1M_Hall_Soother,ConsL1M_Hall_Fruity,ConsL1M_Hall_XS,ConsL1M_Hall_Yellow,ConsL1M_Hall_Green,ConsL1M_Hall_LightBlue,ConsL1M_Hall_HoneyFlav,ConsL1M_Hall_HoneyLemFlav,ConsL1M_Hall_LemonFlavour,ConsL1M_Hall_White,ConsL1M_HartBeat,ConsL1M_KopokoCappuccino,ConsL1M_KopikoSugarFree,ConsL1M_Ole,ConsL1M_Clorets,ConsL1W_Kopipo_org,ConsL1W_Hall_Cool,ConsL1W_Hall_Soother,ConsL1W_Hall_Fruity,ConsL1W_Hall_XS,ConsL1W_Hall_Yellow,ConsL1W_Hall_Green,ConsL1W_Hall_LightBlue,ConsL1W_Hall_HoneyFlav,ConsL1W_Hall_HoneyLemFlav,ConsL1W_Hall_LemonFlavour,ConsL1W_Hall_White,ConsL1W_HartBeat,ConsL1W_KopokoCappuccino,ConsL1W_KopikoSugarFree,ConsL1W_Ole,ConsL1W_Clorets,IntToBuy_Kopipo_org,IntToBuy_Hall_Cool,IntToBuy_Hall_Soother,IntToBuy_Hall_Fruity,IntToBuy_Hall_XS,IntToBuy_Hall_Yellow,IntToBuy_Hall_Green,IntToBuy_Hall_LightBlue,IntToBuy_Hall_HoneyFlav,IntToBuy_Hall_HoneyLemFlav,IntToBuy_Hall_LemonFlavour,IntToBuy_Hall_White,IntToBuy_HartBeat,IntToBuy_KopokoCappuccino,IntToBuy_KopikoSugarFree,IntToBuy_Ole,IntToBuy_Clorets,Bumo,PreBumo,NetBrTom,BrSpontNet_KopikoNet,BrSpontNet_HallsNet,BrSpontNet_HartBeatNet,BrSpontNet_OleNet,BrSpontNet_CloretsNet,BrAidNet_KopikoNet,BrAidNet_HallsNet,BrAidNet_HartBeatNet,BrAidNet_OleNet,BrAidNet_CloretsNet,NetAdTom,AdSpontNet_KopikoNet,AdSpontNet_HallsNet,AdSpontNet_HartBeatNet,AdSpontNet_OleNet,AdSpontNet_CloretsNet,AdAidNet_KopikoNet,AdAidNet_HallsNet,AdAidNet_HartBeatNet,AdAidNet_OleNet,AdAidNet_CloretsNet,EverConsNet_KopikoNet,EverConsNet_HallsNet,EverConsNet_HartBeatNet,EverConsNet_OleNet,EverConsNet_CloretsNet,ConsL3MNet_KopikoNet,ConsL3MNet_HallsNet,ConsL3MNet_HartBeatNet,ConsL3MNet_OleNet,ConsL3MNet_CloretsNet,ConsL1MNet_KopikoNet,ConsL1MNet_HallsNet,ConsL1MNet_HartBeatNet,ConsL1MNet_OleNet,ConsL1MNet_CloretsNet,ConsL1WNet_KopikoNet,ConsL1WNet_HallsNet,ConsL1WNet_HartBeatNet,ConsL1WNet_OleNet,ConsL1WNet_CloretsNet,IntBuyNetNet_KopikoNet,IntBuyNetNet_HallsNet,IntBuyNetNet_HartBeatNet,IntBuyNetNet_OleNet,IntBuyNetNet_CloretsNet,Bumo_Net,PreBumo_Net,r63_1_1,r63_1_2,r63_1_3,r63_1_4,r63_1_5,r63_1_6,r63_1_7,r63_1_8,r63_1_9,r63_1_10,r63_1_11,r63_1_12,r63_1_13,r63_2,r64_1,r64_2,r64_3,r64_4,r64_5,r64_6,r64_7,r67_1_1,r67_1_2,r67_1_3,r67_1_4,r67_1_5,r67_1_6,r67_1_7,r68_2a_1,r68_2a_2,r68_2a_3,r68_2a_4,r68_2a_5,r50h1_1,r50h1_2,r50h1_3,r50h1_5,r50h1_6,r50h1_7,r50h1_8,r50h1_12,r50h1_13,r50h1_14,r50h1_15,r50h2_1,r50h2_2,r50h2_3,r50h2_5,r50h2_12,r50h2_13,r50h2_14,r50h2_15,r50h3_1,r50h3_2,r50h3_3,r50h3_5,r50h3_12,r50h3_13,r50h3_14,r50h3_15,r50h4_1,r50h4_2,r50h4_3,r50h4_5,r50h4_12,r50h4_13,r50h4_14,r50h4_15,r50h5_1,r50h5_2,r50h5_3,r50h5_5,r50h5_12,r50h5_13,r50h5_14,r50h5_15) " + "VALUES('" + UserID + "','" + Country + "','" + SurveyID + "','" + Gender + "','" + MaritalStatus + "','" + SurveyPeriod + "','" + Weight + "','" + Location + "','" + AgeGroup + "','" + Education + "','" + Occupation + "','" + PersonalInc_BKK + "','" + HHIncome_BKK + "','" + PersonalInc_UPC + "','" + HHIncome_UPC + "','" + Period + "','" + S5_1 + "','" + S5_2 + "','" + S5_3 + "','" + S5_4 + "','" + S5_5 + "','" + S5_6 + "','" + S5_7 + "','" + S5_8 + "','" + S5_9 + "','" + S5_10 + "','" + S5_11 + "','" + S5_12 + "','" + S5_13 + "','" + S5_14 + "','" + S5_15 + "','" + S5_16 + "','" + S5_17 + "','" + S5_18 + "','" + S5_19 + "','" + S6 + "','" + S7_1 + "','" + S7_2 + "','" + S7_3 + "','" + S7_4 + "','" + S7_5 + "','" + S7_6 + "','" + S7_7 + "','" + S7_8 + "','" + S7_9 + "','" + S7_10 + "','" + S7_11 + "','" + S7_12 + "','" + S7_13 + "','" + S7_14 + "','" + S7_15 + "','" + S7_16 + "','" + S7_17 + "','" + S7_18 + "','" + S7_19 + "','" + S8 + "','" + S10_2 + "','" + S10_3 + "','" + S11_2 + "','" + S11_3 + "','" + BrTom + "','" + BrSpont_Kopipo_org + "','" + BrSpont_Hall_Cool + "','" + BrSpont_Hall_Soother + "','" + BrSpont_Hall_Fruity + "','" + BrSpont_Hall_XS + "','" + BrSpont_Hall_Yellow + "','" + BrSpont_Hall_Green + "','" + BrSpont_Hall_LightBlue + "','" + BrSpont_Hall_HoneyFlav + "','" + BrSpont_Hall_HoneyLemFlav + "','" + BrSpont_Hall_LemonFlavour + "','" + BrSpont_Hall_White + "','" + BrSpont_HartBeat + "','" + BrSpont_KopokoCappuccino + "','" + BrSpont_KopikoSugarFree + "','" + BrSpont_Ole + "','" + BrSpont_Clorets + "','" + BrAid_Kopipo_org + "','" + BrAid_Hall_Cool + "','" + BrAid_Hall_Soother + "','" + BrAid_Hall_Fruity + "','" + BrAid_Hall_XS + "','" + BrAid_Hall_Yellow + "','" + BrAid_Hall_Green + "','" + BrAid_Hall_LightBlue + "','" + BrAid_Hall_HoneyFlav + "','" + BrAid_Hall_HoneyLemFlav + "','" + BrAid_Hall_LemonFlavour + "','" + BrAid_Hall_White + "','" + BrAid_HartBeat + "','" + BrAid_KopokoCappuccino + "','" + BrAid_KopikoSugarFree + "','" + BrAid_Ole + "','" + BrAid_Clorets + "','" + AdTom + "','" + AdSpont_Kopipo_org + "','" + AdSpont_Hall_Cool + "','" + AdSpont_Hall_Soother + "','" + AdSpont_Hall_Fruity + "','" + AdSpont_Hall_XS + "','" + AdSpont_Hall_Yellow + "','" + AdSpont_Hall_Green + "','" + AdSpont_Hall_LightBlue + "','" + AdSpont_Hall_HoneyFlav + "','" + AdSpont_Hall_HoneyLemFlav + "','" + AdSpont_Hall_LemonFlavour + "','" + AdSpont_Hall_White + "','" + AdSpont_HartBeat + "','" + AdSpont_KopokoCappuccino + "','" + AdSpont_KopikoSugarFree + "','" + AdSpont_Ole + "','" + AdSpont_Clorets + "','" + AdAided_Kopipo_org + "','" + AdAided_Hall_Cool + "','" + AdAided_Hall_Soother + "','" + AdAided_Hall_Fruity + "','" + AdAided_Hall_XS + "','" + AdAided_Hall_Yellow + "','" + AdAided_Hall_Green + "','" + AdAided_Hall_LightBlue + "','" + AdAided_Hall_HoneyFlav + "','" + AdAided_Hall_HoneyLemFlav + "','" + AdAided_Hall_LemonFlavour + "','" + AdAided_Hall_White + "','" + AdAided_HartBeat + "','" + AdAided_KopokoCappuccino + "','" + AdAided_KopikoSugarFree + "','" + AdAided_Ole + "','" + AdAided_Clorets + "','" + EverCons_Kopipo_org + "','" + EverCons_Hall_Cool + "','" + EverCons_Hall_Soother + "','" + EverCons_Hall_Fruity + "','" + EverCons_Hall_XS + "','" + EverCons_Hall_Yellow + "','" + EverCons_Hall_Green + "','" + EverCons_Hall_LightBlue + "','" + EverCons_Hall_HoneyFlav + "','" + EverCons_Hall_HoneyLemFlav + "','" + EverCons_Hall_LemonFlavour + "','" + EverCons_Hall_White + "','" + EverCons_HartBeat + "','" + EverCons_KopokoCappuccino + "','" + EverCons_KopikoSugarFree + "','" + EverCons_Ole + "','" + EverCons_Clorets + "','" + ConsL3M_Kopipo_org + "','" + ConsL3M_Hall_Cool + "','" + ConsL3M_Hall_Soother + "','" + ConsL3M_Hall_Fruity + "','" + ConsL3M_Hall_XS + "','" + ConsL3M_Hall_Yellow + "','" + ConsL3M_Hall_Green + "','" + ConsL3M_Hall_LightBlue + "','" + ConsL3M_Hall_HoneyFlav + "','" + ConsL3M_Hall_HoneyLemFlav + "','" + ConsL3M_Hall_LemonFlavour + "','" + ConsL3M_Hall_White + "','" + ConsL3M_HartBeat + "','" + ConsL3M_KopokoCappuccino + "','" + ConsL3M_KopikoSugarFree + "','" + ConsL3M_Ole + "','" + ConsL3M_Clorets + "','" + ConsL1M_Kopipo_org + "','" + ConsL1M_Hall_Cool + "','" + ConsL1M_Hall_Soother + "','" + ConsL1M_Hall_Fruity + "','" + ConsL1M_Hall_XS + "','" + ConsL1M_Hall_Yellow + "','" + ConsL1M_Hall_Green + "','" + ConsL1M_Hall_LightBlue + "','" + ConsL1M_Hall_HoneyFlav + "','" + ConsL1M_Hall_HoneyLemFlav + "','" + ConsL1M_Hall_LemonFlavour + "','" + ConsL1M_Hall_White + "','" + ConsL1M_HartBeat + "','" + ConsL1M_KopokoCappuccino + "','" + ConsL1M_KopikoSugarFree + "','" + ConsL1M_Ole + "','" + ConsL1M_Clorets + "','" + ConsL1W_Kopipo_org + "','" + ConsL1W_Hall_Cool + "','" + ConsL1W_Hall_Soother + "','" + ConsL1W_Hall_Fruity + "','" + ConsL1W_Hall_XS + "','" + ConsL1W_Hall_Yellow + "','" + ConsL1W_Hall_Green + "','" + ConsL1W_Hall_LightBlue + "','" + ConsL1W_Hall_HoneyFlav + "','" + ConsL1W_Hall_HoneyLemFlav + "','" + ConsL1W_Hall_LemonFlavour + "','" + ConsL1W_Hall_White + "','" + ConsL1W_HartBeat + "','" + ConsL1W_KopokoCappuccino + "','" + ConsL1W_KopikoSugarFree + "','" + ConsL1W_Ole + "','" + ConsL1W_Clorets + "','" + IntToBuy_Kopipo_org + "','" + IntToBuy_Hall_Cool + "','" + IntToBuy_Hall_Soother + "','" + IntToBuy_Hall_Fruity + "','" + IntToBuy_Hall_XS + "','" + IntToBuy_Hall_Yellow + "','" + IntToBuy_Hall_Green + "','" + IntToBuy_Hall_LightBlue + "','" + IntToBuy_Hall_HoneyFlav + "','" + IntToBuy_Hall_HoneyLemFlav + "','" + IntToBuy_Hall_LemonFlavour + "','" + IntToBuy_Hall_White + "','" + IntToBuy_HartBeat + "','" + IntToBuy_KopokoCappuccino + "','" + IntToBuy_KopikoSugarFree + "','" + IntToBuy_Ole + "','" + IntToBuy_Clorets + "','" + Bumo + "','" + PreBumo + "','" + NetBrTom + "','" + BrSpontNet_KopikoNet + "','" + BrSpontNet_HallsNet + "','" + BrSpontNet_HartBeatNet + "','" + BrSpontNet_OleNet + "','" + BrSpontNet_CloretsNet + "','" + BrAidNet_KopikoNet + "','" + BrAidNet_HallsNet + "','" + BrAidNet_HartBeatNet + "','" + BrAidNet_OleNet + "','" + BrAidNet_CloretsNet + "','" + NetAdTom + "','" + AdSpontNet_KopikoNet + "','" + AdSpontNet_HallsNet + "','" + AdSpontNet_HartBeatNet + "','" + AdSpontNet_OleNet + "','" + AdSpontNet_CloretsNet + "','" + AdAidNet_KopikoNet + "','" + AdAidNet_HallsNet + "','" + AdAidNet_HartBeatNet + "','" + AdAidNet_OleNet + "','" + AdAidNet_CloretsNet + "','" + EverConsNet_KopikoNet + "','" + EverConsNet_HallsNet + "','" + EverConsNet_HartBeatNet + "','" + EverConsNet_OleNet + "','" + EverConsNet_CloretsNet + "','" + ConsL3MNet_KopikoNet + "','" + ConsL3MNet_HallsNet + "','" + ConsL3MNet_HartBeatNet + "','" + ConsL3MNet_OleNet + "','" + ConsL3MNet_CloretsNet + "','" + ConsL1MNet_KopikoNet + "','" + ConsL1MNet_HallsNet + "','" + ConsL1MNet_HartBeatNet + "','" + ConsL1MNet_OleNet + "','" + ConsL1MNet_CloretsNet + "','" + ConsL1WNet_KopikoNet + "','" + ConsL1WNet_HallsNet + "','" + ConsL1WNet_HartBeatNet + "','" + ConsL1WNet_OleNet + "','" + ConsL1WNet_CloretsNet + "','" + IntBuyNetNet_KopikoNet + "','" + IntBuyNetNet_HallsNet + "','" + IntBuyNetNet_HartBeatNet + "','" + IntBuyNetNet_OleNet + "','" + IntBuyNetNet_CloretsNet + "','" + Bumo_Net + "','" + PreBumo_Net + "','" + r63_1_1 + "','" + r63_1_2 + "','" + r63_1_3 + "','" + r63_1_4 + "','" + r63_1_5 + "','" + r63_1_6 + "','" + r63_1_7 + "','" + r63_1_8 + "','" + r63_1_9 + "','" + r63_1_10 + "','" + r63_1_11 + "','" + r63_1_12 + "','" + r63_1_13 + "','" + r63_2 + "','" + r64_1 + "','" + r64_2 + "','" + r64_3 + "','" + r64_4 + "','" + r64_5 + "','" + r64_6 + "','" + r64_7 + "','" + r67_1_1 + "','" + r67_1_2 + "','" + r67_1_3 + "','" + r67_1_4 + "','" + r67_1_5 + "','" + r67_1_6 + "','" + r67_1_7 + "','" + r68_2a_1 + "','" + r68_2a_2 + "','" + r68_2a_3 + "','" + r68_2a_4 + "','" + r68_2a_5 + "','" + r50h1_1 + "','" + r50h1_2 + "','" + r50h1_3 + "','" + r50h1_5 + "','" + r50h1_6 + "','" + r50h1_7 + "','" + r50h1_8 + "','" + r50h1_12 + "','" + r50h1_13 + "','" + r50h1_14 + "','" + r50h1_15 + "','" + r50h2_1 + "','" + r50h2_2 + "','" + r50h2_3 + "','" + r50h2_5 + "','" + r50h2_12 + "','" + r50h2_13 + "','" + r50h2_14 + "','" + r50h2_15 + "','" + r50h3_1 + "','" + r50h3_2 + "','" + r50h3_3 + "','" + r50h3_5 + "','" + r50h3_12 + "','" + r50h3_13 + "','" + r50h3_14 + "','" + r50h3_15 + "','" + r50h4_1 + "','" + r50h4_2 + "','" + r50h4_3 + "','" + r50h4_5 + "','" + r50h4_12 + "','" + r50h4_13 + "','" + r50h4_14 + "','" + r50h4_15 + "','" + r50h5_1 + "','" + r50h5_2 + "','" + r50h5_3 + "','" + r50h5_5 + "','" + r50h5_12 + "','" + r50h5_13 + "','" + r50h5_14 + "','" + r50h5_15 + "' )", connection);
            try
            {
                cmd.ExecuteNonQuery();
                Console.WriteLine("Data inserted successfully");
                i = 1;

            }
            catch (Exception ex)
            {

                connection.Close();
                i = 0;
                Console.WriteLine("Exception occured" + ex.ToString());
                Console.WriteLine("Exception occured");

                Console.ReadLine();
            }
            connection.Close();
        }


    }
}
