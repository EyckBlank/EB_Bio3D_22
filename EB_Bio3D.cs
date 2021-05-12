///////////////////////////////////////////////////////////////////////////////////////////////////
// EB_Bio3D
// zur Biologischen Umrechnung von Dosisverteilungen
//
// zuerst einen neuen Plan anlegen (Kopie des alten Planes mit Dosismatrix, statt -RA -BIO)
// Felder werden nicht gebraucht
// dann welcher Dosispunkt geh�rt zu welchem Volumen
// Dann umrechnen der Dosismatrix strukturweise
//
// geplant 20201229:
// vor globaler String Declaration alles nach hinten
// alle DoseAtVolumes etc weg
// wie Umgang mit substrings
// Umgang mit Organ None etc
// 
// created at 21.12.2020 by Eyck Blank
// modified at 23.2.2021
//
// EB_Bio3D
// korrigiert
// 2 Vermutungen
// 1. doppelte EQD2 Berechnung m�glich
// 2. GTV EQD2 kann �bert�ncht werden
//
// 1. Vermutung
// Es wurde zus�tzliche ebuffer Array eingef�hrt
// damit wird ebuffer immer aus originalem dbuffer berechnet
// Doppeltberechnung nicht mehr m�glich
//
// 2. Vermutung
// EQD2 Berechnung f�r GRTV ans Ende
// damit dieses Primat hat und nicht mehr �berget�ncht werden kann. 
//
// dazu
// GTV besonders detektiert
// Liste aller Strukturen ausgelesen
// Daraus SubListe der GTV's gefiltert
//
// Bemerkung
// es gibt zu viele GTV Namen
// Filterung mit ersten drei Buchstaben GTV
// (es k�nnen ja mehrere GTV's sein)
// Test
// sieht gut aus 
// Proberechnung EQD2 der Lungen stimmen
// 
// 31.3.2021   1.0.0.3
// GTV Erkennung   line commentiert 409-411
// fraktDose statt 2.0 in Formel line 465
///////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using System.Linq;
using System.Windows;
using System.Collections.Generic;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.0.0.3")]
[assembly: AssemblyFileVersion("1.0.0.3")]
[assembly: AssemblyInformationalVersion("1.03")]

// TODO: Uncomment the following line if the script requires write access.
   [assembly: ESAPIScript(IsWriteable = true)]


namespace VMS.TPS
{
 
    public class Script
    {
        public Script()
        {
        }

        double BioDose;

        const string BODY_ID1 = "BODY";
        const string BODY_ID2 = "Body";

        const string HIRN_ID = "Hirn";
        const string HIRNSTAMM_ID = "Hirnstamm";
        const string HIPPOCAMPUS_R_ID ="Hippocampus"+" "+"re";
        const string HIPPOCAMPUS_L_ID ="Hippocampus"+" "+"li";

        const string COCLEA_R_ID = "Coclea"+" "+"re";
        const string COCLEA_L_ID = "Coclea"+" "+"li";

        const string CHIASMA_ID = "Chiasma";
        const string SEHNERV_R_ID = "Sehnerv"+" "+"re";
        const string SEHNERV_L_ID = "Sehnerv"+" "+"li";

        const string LINSE_R_ID = "Linse"+" "+"re";
        const string LINSE_L_ID = "Linse"+" "+"li";
        
        const string PAROTIS_R_ID = "Parotis"+" "+"re";
        const string PAROTIS_L_ID = "Parotis"+" "+"li";

        const string SUBMANDI_R_ID = "Submandi"+" "+"re";
        const string SUBMANDI_L_ID = "Submandi"+" "+"li";
        
        const string MANDIBULA_ID = "Mandibula";
        const string CONSTRICTOR_ID = "Constrictor";

        const string MYELON_ID = "Myelon";

        const string PLEXUS_R_ID = "Plexus"+" "+"re";
        const string PLEXUS_L_ID = "Plexus"+" "+"li";
            
        const string LUNGE_R_ID = "Lunge"+" "+"re";
        const string LUNGE_L_ID = "Lunge"+" "+"li";
            
        const string HERZ_ID = "Herz";
        const string RIVA_ID = "Riva";

        const string BRUST_R_ID = "Brust"+" "+"re";
        const string BRUST_L_ID = "Brust"+" "+"li";

        const string OESOPHAGUS_ID = "Oesophagus";
        const string LEBER_ID = "Leber";

        const string NIERE_R_ID = "Niere"+" "+"re";
        const string NIERE_L_ID = "Niere"+" "+"li";

        const string DARM_ID = "Darm";
        const string REKTUM_ID = "Rektum";
        const string BLASE_ID = "Blase";

        const string GTV_ID = "GTV";

        const string SCRIPT_NAME = "Bio IsodosenPlan Script";


        //---------------------------------------------------------------------------------------------  
        // public void Execute(ScriptContext context, Window window)   // if  a window should be shown
        public void Execute(ScriptContext context)
        {
            string PatLName = "";
            string PatFName = "";
            string PatID = "";
            string sBody = "";

            string PlanPTV = "";
            string PlanPrescr = "";

            double GD;
            double ED;
            double N;   //GD=ED*N
            double PI;  //prescr.isodose

            List<string> gtv = new List<string>();
            List<Structure> gtvls = new List<Structure>();
 
            // a/b values,  later into excel file
            double abBody = 2;
            double abHirn = 2;
            double abHirnstamm = 2;
            double abHippocampus = 2;
            double abCoclea = 2;
            double abChiasma = 2;
            double abSehnerv = 2;
            double abLinse = 2;
            double abParotis = 2;
            double abSubmandi = 2;
            double abMandibula = 2;
            double abConstrictor = 2;
            double abMyelon = 2;
            double abPlexus = 2;
            double abLunge = 2;
            double abHerz = 2;
            double abRiva = 2;
            double abBrust = 2;
            double abOesophagus = 2;
            double abLeber = 2;
            double abNiere = 2;
            double abDarm = 2;
            double abRektum = 2;
            double abBlase = 2;
            double abGtv = 8;
            

            // patient and plan context
            Patient Pat = context.Patient;
            PatLName = context.Patient.LastName.ToString();
            PatFName = context.Patient.FirstName.ToString();
            PatID = context.Patient.Id.ToString();

            context.Patient.BeginModifications();

            Course eCourse = context.Course;
            
            ExternalPlanSetup plan = context.ExternalPlanSetup;
                            
            if (context.Patient == null || context.StructureSet == null)
            {
                MessageBox.Show("Please load a patient, 3D image, and structure set before running this script.", SCRIPT_NAME, MessageBoxButton.OKCancel, MessageBoxImage.Exclamation);
                return;
            }
            StructureSet ss = context.StructureSet;
            
            foreach(Structure s in ss.Structures)
            {
                if (s.Id.Length>3)
                {
                    if (s.Id == BODY_ID1 | s.Id == BODY_ID2)  // because upper-/lowercase
                    {
                        sBody = s.Id;
                    }
                    if (s.Id.Substring(0,3) == "GTV")
                    {
                        gtv.Add(s.Id);
                        gtvls.Add(s);
                    }
                }
            }
            
            var ePlan = eCourse.AddExternalPlanSetupAsVerificationPlan(ss, plan);
            
            // ePlanName
            String planName = plan.Id;
            int startIndex = 0;
            int length = 10;
            string zName = planName;
            int zLength = zName.Length;
            if (zLength>10)
            {
                length = 10;
            }
            else 
            {
                length = zLength;
            }
            string ePlanName = planName.Substring(startIndex, length) + "-E2";

            ePlan.Id = ePlanName;
                                                                                
            MessageBox.Show("VeriPlan angelegt", SCRIPT_NAME, MessageBoxButton.OKCancel, MessageBoxImage.Exclamation);
            

            if (plan == null)
                return;

            int Frakt = plan.NumberOfFractions.Value;

            PlanPTV = plan.TargetVolumeID;

            GD = plan.TotalDose.Dose;
            ED = plan.DosePerFraction.Dose;
            N = (double)plan.NumberOfFractions.Value;
            PI = plan.TreatmentPercentage * 100;

            string sGD = GD.ToString("F2") + plan.TotalDose.Unit;
            string sED = ED.ToString("F2");
            string sN = N.ToString("F0");
            string sPI = PI.ToString("F2");

            PlanPrescr = sED + " x " + sN + " = " + sGD + " (" + sPI + "%)";

            //=========================
            // Find the  structures
            //=========================
            
            // find Body 
            Structure body = ss.Structures.FirstOrDefault(x => x.Id == sBody);  // wegen Gro�- oder Kleinschreibung
           
            // find Hirn (brain)
            Structure hirn = ss.Structures.FirstOrDefault(x => x.Id == HIRN_ID);
           
            // find Hirnstamm (brainstem)
            Structure hirnstamm = ss.Structures.FirstOrDefault(x => x.Id == HIRNSTAMM_ID);
           
            // find Hippocampus_re (hippocampus right)
            Structure hippocampus_re = ss.Structures.FirstOrDefault(x => x.Id == HIPPOCAMPUS_R_ID);
           
            // find Hippocampus_li (hippocampus left)
            Structure hippocampus_li = ss.Structures.FirstOrDefault(x => x.Id == HIPPOCAMPUS_L_ID);
           
            // find Coclea_re (coclea right)
            Structure coclea_re = ss.Structures.FirstOrDefault(x => x.Id == COCLEA_R_ID);
           
            // find Coclea_li (coclea left)
            Structure coclea_li = ss.Structures.FirstOrDefault(x => x.Id == COCLEA_L_ID);
           
            // find Chiasma (chiasma opticus)
            Structure chiasma = ss.Structures.FirstOrDefault(x => x.Id == CHIASMA_ID);
           
            // find Sehnerv_re (nervus opticus right)
            Structure sehnerv_re = ss.Structures.FirstOrDefault(x => x.Id == SEHNERV_R_ID);
           
            // find Sehnerv_li (nervus opticus lrft)
            Structure sehnerv_li = ss.Structures.FirstOrDefault(x => x.Id == SEHNERV_L_ID);
           
            // find Linse_re (lens right)
            Structure linse_re = ss.Structures.FirstOrDefault(x => x.Id == LINSE_R_ID);
           
            // find Linse_li (lens left)
            Structure linse_li = ss.Structures.FirstOrDefault(x => x.Id == LINSE_L_ID);
          
            // find Parotis_re (gland parotis right)
            Structure parotis_re = ss.Structures.FirstOrDefault(x => x.Id == PAROTIS_R_ID);
          
           // find Parotis_li (gland parotis left)
            Structure parotis_li = ss.Structures.FirstOrDefault(x => x.Id == PAROTIS_L_ID);
           
            // find Submandi_re (gland submand right)
            Structure submandi_re = ss.Structures.FirstOrDefault(x => x.Id == SUBMANDI_R_ID);
          
            // find Submandi_li (gland (submand left)
            Structure submandi_li = ss.Structures.FirstOrDefault(x => x.Id == SUBMANDI_L_ID);
           
            // find Mandibula (mandibula)
            Structure mandibula = ss.Structures.FirstOrDefault(x => x.Id == MANDIBULA_ID);
           
            // find Constrictoor (muscule constrictor pharyngialis))
            Structure constrictor = ss.Structures.FirstOrDefault(x => x.Id == CONSTRICTOR_ID);
           
            // find Myelon (myelon)
            Structure myelon = ss.Structures.FirstOrDefault(x => x.Id == MYELON_ID);
                                  
            // find Plexus_re (nervus plexus right)
            Structure plexus_re = ss.Structures.FirstOrDefault(x => x.Id == PLEXUS_R_ID);
          
            // find Plexus_li (nervus plexus left)
            Structure plexus_li = ss.Structures.FirstOrDefault(x => x.Id == PLEXUS_L_ID);
          
            // find Lunge_re (lung right)
            Structure lunge_re = ss.Structures.FirstOrDefault(x => x.Id == LUNGE_R_ID);
           
            // find Lunge_li (lung right)
            Structure lunge_li = ss.Structures.FirstOrDefault(x => x.Id == LUNGE_L_ID);
           
            // find Herz (heart)
            Structure herz = ss.Structures.FirstOrDefault(x => x.Id == HERZ_ID);
          
            // find Riva (v. riva)
            Structure riva = ss.Structures.FirstOrDefault(x => x.Id == RIVA_ID);
          
            // find Brust_re (breast right)
            Structure brust_re = ss.Structures.FirstOrDefault(x => x.Id == BRUST_R_ID);
          
            // find Brust_li (breast left)
            Structure brust_li = ss.Structures.FirstOrDefault(x => x.Id == BRUST_L_ID);
          
            // find Oesophagus (oesophagus)
            Structure oesophagus = ss.Structures.FirstOrDefault(x => x.Id == OESOPHAGUS_ID);
          
            // find Leber (liver)
            Structure leber = ss.Structures.FirstOrDefault(x => x.Id == LEBER_ID);
           
            // find Niere_re (kidney right)
            Structure niere_re = ss.Structures.FirstOrDefault(x => x.Id == NIERE_R_ID);
           
            // find Niere_li (kidney left)
            Structure niere_li = ss.Structures.FirstOrDefault(x => x.Id == NIERE_L_ID);
                      
            // find Darm (bowel)
            Structure darm = ss.Structures.FirstOrDefault(x => x.Id == DARM_ID);
         
            // find Rektum (rectum)
            Structure rektum = ss.Structures.FirstOrDefault(x => x.Id == REKTUM_ID);
           
            // find Blase (bladder)
            Structure blase = ss.Structures.FirstOrDefault(x => x.Id == BLASE_ID);

            // find GTV
            // how can I find all GTVs of structureset ?
            // Structure gtv1 = ss.Structures.FirstOrDefault(x => x.Id.Substring(0, 3) == GTV_ID);
           

            //=======================================
            // calculate the transformation parameters
            //=======================================
          
            var dose = plan.Dose;
                            
            int eX = 0;  // for debugger
            int eY = 0;
            int eZ = 0;

            double erX = 0;  // for debugger
            double erY = 0;
            double erZ = 0;

            int dSizeIX = dose.XSize;
            int dSizeIY = dose.YSize;
            int dSizeIZ = dose.YSize;

            double doX = dose.Origin.x;
            double doY = dose.Origin.y;
            double doZ = dose.Origin.z;

            double dsizeX = dose.XSize;
            double dsizeY = dose.YSize;
            double dsizeZ = dose.ZSize;

            double dresX = dose.XRes;
            double dresY = dose.YRes;
            double dresZ = dose.ZRes;

            int[,] zbuffer    = new int[dose.XSize, dose.YSize];
            double[,] dbuffer = new double[dose.XSize, dose.YSize];
            double[,] ebuffer = new double[dose.XSize, dose.YSize];
            double[] dbu = new double[dose.ZSize];
            int[] dbi    = new int[dose.ZSize];
            
            double[] xm = new double[dose.XSize];
            double[] ym = new double[dose.YSize];
            double[] zm = new double[dose.ZSize];

            EvaluationDose eDose = ePlan.CreateEvaluationDose();
            ePlan.CopyEvaluationDose(dose);         // to fit dimensions of dose matrices  dose und eDose !
            double fraktDose = plan.DosePerFraction.Dose;
            int fractNumber = (int)plan.NumberOfFractions;
            DoseValue ee = new DoseValue(fraktDose, DoseValue.DoseUnit.Gy);  // prepare fraction dose for new ePlan 
            ePlan.SetPrescription(fractNumber, ee, 1.0);               // here set fraction number and fraction dose of ePlan in ARIA 

            // whole 3D matrix will be scanned
            VVector VV = new VVector(0,0,0);
            if (dose != null)
            {
                for (int zi = 0; zi < dose.ZSize; zi++)
                {
                    zm[zi] = zi * dresZ + doZ;
                    dose.GetVoxels(zi, zbuffer);   // here read a dose layer

                    dbi[zi] = zbuffer[90, 54];     // only for test by debuging, how look the dose values
                                          
                    for (int yi = 0; yi < dose.YSize; yi++)
                    {
                        ym[yi] = yi * dresY + doY;
                            
                        for (int xi = 0; xi < dose.XSize; xi++)
                        {
                            xm[xi] = xi * dresX + doX;  // calculate carthesian points of matrix points for PointInsideSegment
                            VV[0] =  xm[xi];
                            VV[1] =  ym[yi];
                            VV[2] =  zm[zi];

                            // IsPointInsideSegment for all entities 
                            // (long list)
                            // EQD2 conversion by a/b values
                            // 
                            // imageMatrix und doseMatrix have different dimensions
                            // IsPointInsideSegment must have carthesian coordinates
                            // x, y, z must converted by Xres, Yres, Zres and Origin 

                            dbuffer[xi, yi] = dose.VoxelToDoseValue(zbuffer[xi, yi]).Dose * GD / 100;
                            ebuffer[xi, yi] = 0;

                            // Body
                            if (body.IsPointInsideSegment(VV))
                            {
                                if (body != null)
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abBody) / (2 + abBody) * dbuffer[xi, yi];
                                }
                            }

                            // Hirn (brain)
                            if (hirn != null)
                            {
                                if (hirn.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abHirn) / (2 + abHirn) * dbuffer[xi, yi];
                                }
                            }

                            // Hirnstamm (brainstem)
                            if (hirnstamm != null)
                            {
                                if (hirnstamm.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abHirnstamm) / (2 + abHirnstamm) * dbuffer[xi, yi];
                                }
                            }

                            // Hippocampus_re (hippocampus right)
                            if (hippocampus_re != null)
                            {
                                if (hippocampus_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abHippocampus) / (2 + abHippocampus) * dbuffer[xi, yi];
                                }
                            }
                            // Hippocampus_li (hippocampus left)
                            if (hippocampus_li != null)
                            {
                                if (hippocampus_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abHippocampus) / (2 + abHippocampus) * dbuffer[xi, yi];
                                }
                            }

                            // Coclea_re (coclea right)
                            if (coclea_re != null)
                            {
                                if (coclea_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abCoclea) / (2 + abCoclea) * dbuffer[xi, yi];
                                }
                            }
                            // Coclea_li (coclea left)
                            if (coclea_li != null)
                            {
                                if (coclea_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abCoclea) / (2 + abCoclea) * dbuffer[xi, yi];
                                }
                            }
                                 
                            // Chiasma (chiasma)
                            if (chiasma != null)
                            {                                
                                if (chiasma.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abChiasma) / (2 + abChiasma) * dbuffer[xi, yi];
                                }
                            }

                            // Sehnerv_re (nervus opticus right)
                            if (sehnerv_re != null)
                            {                                
                                if (sehnerv_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abSehnerv) / (2 + abSehnerv) * dbuffer[xi, yi];
                                }
                            }
                            // Sehnerv_li (nervus opticus left)
                            if (sehnerv_li != null)
                            {                                
                                if (sehnerv_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abSehnerv) / (2 + abSehnerv) * dbuffer[xi, yi];
                                }
                            }

                            // Linse_re (lens right)
                            if (linse_re  != null)
                            {                                
                                if (linse_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abLinse) / (2 + abLinse) * dbuffer[xi, yi];
                                }
                            }
                            // Linse_li (lens left)
                            if (linse_li  != null)
                            {                                
                                if (linse_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abLinse) / (2 + abLinse) * dbuffer[xi, yi];
                                }
                            }

                            // Parotis_re (gland parotis right)
                            if (parotis_re  != null)
                            {                                
                                if (parotis_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abParotis) / (2 + abParotis) * dbuffer[xi, yi];
                                }
                            }
                            // Parotis_li (gland parotis left)
                            if (parotis_li  != null)
                            {                                
                                if (parotis_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abParotis) / (2 + abParotis) * dbuffer[xi, yi];
                                }
                            }

                            // Submandi_re (gland submand right)
                            if (submandi_re != null)
                            {                                
                                if (submandi_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abSubmandi) / (2 + abSubmandi) * dbuffer[xi, yi];
                                }
                            }
                            // Submandi_li (gland submand left)
                            if (submandi_li != null)
                            {                                
                                if (submandi_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abSubmandi) / (2 + abSubmandi) * dbuffer[xi, yi];
                                }
                            }

                            // Mandibula (mandibule)
                            if (mandibula != null)
                            {                                
                                if (mandibula.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abMandibula) / (2 + abMandibula) * dbuffer[xi, yi];
                                }
                            }

                            // Constrictor (muscule constrictor pharyngialis)
                            if (constrictor != null)
                            {                                
                                if (constrictor.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abConstrictor) / (2 + abConstrictor) * dbuffer[xi, yi];
                                }
                            }

                            // Myelon (myelon)
                            if (myelon != null)
                            {                                
                                if (myelon.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abMyelon) / (2 + abMyelon) * dbuffer[xi, yi];
                                }
                            }
                         
                            // Plexus_re (nervus plexus right)
                            if (plexus_re != null)
                            {                                
                                if (plexus_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abPlexus) / (2 + abPlexus) * dbuffer[xi, yi];
                                }
                            }
                            // Plexus_li (nervus plexus left)
                            if (plexus_li != null)
                            {                                
                                if (plexus_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abPlexus) / (2 + abPlexus) * dbuffer[xi, yi];
                                }
                            }

                            // Lunge_re (lung right)
                            if (lunge_re != null)
                            {                                
                                if (lunge_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abLunge) / (2 + abLunge) * dbuffer[xi, yi];
                                }
                            }
                            // Lunge_li (lung left)
                            if (lunge_li != null)
                            {                                
                                if (lunge_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abLunge) / (2 + abLunge) * dbuffer[xi, yi];
                                }
                            }

                            // Herz (heart)
                            if (herz != null)
                            {                                
                                if (herz.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abHerz) / (2 + abHerz) * dbuffer[xi, yi];
                                }
                            }

                            // Riva (v. riva)
                            if (riva != null)
                            {                                
                                if (riva.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abRiva) / (2 + abRiva) * dbuffer[xi, yi];
                                }
                            }

                            // Brust_re (breast right)
                            if (brust_re != null)
                            {                                
                                if (brust_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abBrust) / (2 + abBrust) * dbuffer[xi, yi];
                                }
                            }
                            // Brust_li (breast left)
                            if (brust_li != null)
                            {                                
                                if (brust_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abBrust) / (2 + abBrust) * dbuffer[xi, yi];
                                }
                            }

                            // Oesophagus (oesophagus)
                            if (oesophagus != null)
                            {                                
                                if (oesophagus.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abOesophagus) / (2 + abOesophagus) * dbuffer[xi, yi];
                                }
                            }
                                                       
                            // Leber (liver)
                            if (leber != null)
                            {                                
                                if (leber.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abLeber) / (2 + abLeber) * dbuffer[xi, yi];
                                }
                            }

                            // Niere_re (kidney right)
                            if (niere_re != null)
                            {                                
                                if (niere_re.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abNiere) / (2 + abNiere) * dbuffer[xi, yi];
                                }
                            }
                            // Niere_li (kidney left)
                            if (niere_li != null)
                            {                                
                                if (niere_li.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abNiere) / (2 + abNiere) * dbuffer[xi, yi];
                                }
                            }

                            // Darm (bowel / gut)
                            if (darm != null)
                            {                                
                                if (darm.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abDarm) / (2 + abDarm) * dbuffer[xi, yi];
                                }
                            }

                            // Rektum (rectum)
                            if (rektum != null)
                            {                                
                                if (rektum.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abRektum) / (2 + abRektum) * dbuffer[xi, yi];
                                }
                            }

                            // Blase (bladder)
                            if (blase != null)
                            {                                
                                if (blase.IsPointInsideSegment(VV))
                                {
                                    ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abBlase) / (2 + abBlase) * dbuffer[xi, yi];
                                }
                            }
                            
                            // GTV list (gtvls)
                            if (gtvls != null)
                            {   
                                foreach (Structure s1 in gtvls)
                                {
                                    if (s1.IsPointInsideSegment(VV))
                                    {
                                        ebuffer[xi, yi] = (dbuffer[xi, yi]/Frakt + abGtv) / (2 + abGtv) * dbuffer[xi, yi];
                                    }
                                }
                            }

                            // all ebuffer reconvert into relative values
                            ebuffer[xi, yi] = ebuffer[xi, yi] * 100 / GD;
                            if (ebuffer[xi, yi] < 0)
                            {
                                ebuffer[xi, yi] = 0;
                            }
                           
                            DoseValue eee = new DoseValue(ebuffer[xi, yi], DoseValue.DoseUnit.Percent);
                            
                            zbuffer[xi, yi] = (int)eDose.DoseValueToVoxel(eee) ; 
                                                                                                         
                        }
                    }
                    
                    // dbu[zi] = ebuffer[90, 54]; for Excel debugging
                    dbi[zi] = zbuffer[90, 54];
                    
                    // for monitoring the matrix dimensions
                    eX = eDose.XSize;
                    eY = eDose.YSize;
                    eZ = eDose.ZSize;
                    erX = eDose.XRes;
                    erY = eDose.YRes;
                    erZ = eDose.ZRes;   

                    eDose.SetVoxels(zi, zbuffer);
                    
                }

               //----------------------------------------------------------------------------------
               // call of Debugging Excel
               //----------------------------------------------------------------------------------
                /*
                UpdateExcel(erX.ToString(), erY.ToString(), erZ.ToString(), doX.ToString(), doY.ToString(), doZ.ToString(),
                    eX.ToString(), eY.ToString(), eZ.ToString(), 
                    xm[0].ToString(), ym[0].ToString(), zm[0].ToString(), 
                    xm[182].ToString(), ym[107].ToString(), zm[194].ToString(), 
                    dbi[0].ToString(), dbi[1].ToString(), dbi[2].ToString(), dbi[3].ToString(), dbi[4].ToString(), dbi[5].ToString(), dbi[6].ToString(), dbi[7].ToString(), dbi[8].ToString(), dbi[9].ToString(), 
                    dbi[10].ToString(), dbi[11].ToString(), dbi[12].ToString(), dbi[13].ToString(), dbi[14].ToString(), dbi[15].ToString(), dbi[16].ToString(), dbi[17].ToString(), dbi[18].ToString(), dbi[19].ToString(), 
                    dbi[20].ToString(), dbi[21].ToString(), dbi[22].ToString(), dbi[23].ToString(), dbi[24].ToString(), dbi[25].ToString(), dbi[26].ToString(), dbi[27].ToString(), dbi[28].ToString(), dbi[29].ToString(), 
                    dbi[30].ToString(), dbi[31].ToString(), dbi[32].ToString(), dbi[33].ToString(), dbi[34].ToString(), dbi[35].ToString(), dbi[36].ToString(), dbi[37].ToString(), dbi[38].ToString(), dbi[39].ToString(), 
                    dbi[40].ToString(), dbi[41].ToString(), dbi[42].ToString(), dbi[43].ToString(), dbi[44].ToString(), dbi[45].ToString(), dbi[46].ToString(), dbi[47].ToString(), dbi[48].ToString(), dbi[49].ToString(), 
                    dbi[50].ToString(), dbi[51].ToString(), dbi[52].ToString(), dbi[53].ToString(), dbi[54].ToString(), dbi[55].ToString(), dbi[56].ToString(), dbi[57].ToString(), dbi[58].ToString(), dbi[59].ToString(), 
                    dbi[60].ToString(), dbi[61].ToString(), dbi[62].ToString(), dbi[63].ToString(), dbi[64].ToString(), dbi[65].ToString(), dbi[66].ToString(), dbi[67].ToString(), dbi[68].ToString(), dbi[69].ToString(), 
                    dbi[70].ToString(), dbi[71].ToString(), dbi[72].ToString(), dbi[73].ToString(), dbi[74].ToString(), dbi[75].ToString(), dbi[76].ToString(), dbi[77].ToString(), dbi[78].ToString(), dbi[79].ToString(), 
                    dbi[80].ToString(), dbi[81].ToString(), dbi[82].ToString(), dbi[83].ToString(), dbi[84].ToString(), dbi[85].ToString(), dbi[86].ToString(), dbi[87].ToString(), dbi[88].ToString(), dbi[89].ToString(), 
                    dbi[90].ToString(), dbi[91].ToString(), dbi[92].ToString(), dbi[93].ToString(), dbi[94].ToString(), dbi[95].ToString(), dbi[96].ToString(), dbi[97].ToString(), dbi[98].ToString(), dbi[99].ToString() );
                */
            }
          
        }
        //################################################################################################################  
        //----------------------------------------------------------------------------------
        // Debugging Excel
        //----------------------------------------------------------------------------------
        private void UpdateExcel(string S1, string S2, string S3, string S4, string S5, string S6, 
            string S11, string S12, string S13, 
            string S21, string S22, string S23, 
            string S31, string S32, string S33, 
            string S100, string S101, string S102, string S103, string S104, string S105, string S106, string S107, string S108, string S109, 
            string S110, string S111, string S112, string S113, string S114, string S115, string S116, string S117, string S118, string S119,  
            string S120, string S121, string S122, string S123, string S124, string S125, string S126, string S127, string S128, string S129,  
            string S130, string S131, string S132, string S133, string S134, string S135, string S136, string S137, string S138, string S139,  
            string S140, string S141, string S142, string S143, string S144, string S145, string S146, string S147, string S148, string S149,  
            string S150, string S151, string S152, string S153, string S154, string S155, string S156, string S157, string S158, string S159,  
            string S160, string S161, string S162, string S163, string S164, string S165, string S166, string S167, string S168, string S169, 
            string S170, string S171, string S172, string S173, string S174, string S175, string S176, string S177, string S178, string S179,  
            string S180, string S181, string S182, string S183, string S184, string S185, string S186, string S187, string S188, string S189, 
            string S190, string S191, string S192, string S193, string S194, string S195, string S196, string S197, string S198, string S199)
        {
            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel._Worksheet oSheet = null;

            try
            {
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oWB = oXL.Workbooks.Open("Q:/ESAPI/Plugins/EB_Bio3D/EB_Bio3D.xlsx");
                oSheet = String.IsNullOrEmpty("Tab1") ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets["Tab1"];

                //.................................................................................
                
                // read the row counter in Excel
                string sReihe = oSheet.Cells[1, 1].Value == null ? "-" : oSheet.Cells[1, 1].Value.ToString();
                int iReihe = Convert.ToInt32(sReihe);
                iReihe = iReihe + 1;

                // dresX
                oSheet.Cells[iReihe, 1] = S1;
                oSheet.Cells[iReihe, 2] = S2;
                oSheet.Cells[iReihe, 3] = S3;
                // doX
                oSheet.Cells[iReihe, 4] = S4;
                oSheet.Cells[iReihe, 5] = S5;
                oSheet.Cells[iReihe, 6] = S6;
                // dsizeX
                oSheet.Cells[iReihe, 8] = S11;
                oSheet.Cells[iReihe, 9] = S12;
                oSheet.Cells[iReihe, 10] = S13;
                // xm start
                oSheet.Cells[iReihe, 12] = S21;
                oSheet.Cells[iReihe, 13] = S22;
                oSheet.Cells[iReihe, 14] = S23;
                // xm stop
                oSheet.Cells[iReihe, 15] = S31;
                oSheet.Cells[iReihe, 16] = S32;
                oSheet.Cells[iReihe, 17] = S33;
                // curve data
                oSheet.Cells[10, 4] = S100;
                oSheet.Cells[11, 4] = S101;
                oSheet.Cells[12, 4] = S102;
                oSheet.Cells[13, 4] = S103;
                oSheet.Cells[14, 4] = S104;
                oSheet.Cells[15, 4] = S105;
                oSheet.Cells[16, 4] = S106;
                oSheet.Cells[17, 4] = S107;
                oSheet.Cells[18, 4] = S108;
                oSheet.Cells[19, 4] = S109;

                oSheet.Cells[20, 4] = S110;
                oSheet.Cells[21, 4] = S111;
                oSheet.Cells[22, 4] = S112;
                oSheet.Cells[23, 4] = S113;
                oSheet.Cells[24, 4] = S114;
                oSheet.Cells[25, 4] = S115;
                oSheet.Cells[26, 4] = S116;
                oSheet.Cells[27, 4] = S117;
                oSheet.Cells[28, 4] = S118;
                oSheet.Cells[29, 4] = S119;

                oSheet.Cells[30, 4] = S120;
                oSheet.Cells[31, 4] = S121;
                oSheet.Cells[32, 4] = S122;
                oSheet.Cells[33, 4] = S123;
                oSheet.Cells[34, 4] = S124;
                oSheet.Cells[35, 4] = S125;
                oSheet.Cells[36, 4] = S126;
                oSheet.Cells[37, 4] = S127;
                oSheet.Cells[38, 4] = S128;
                oSheet.Cells[39, 4] = S129;

                oSheet.Cells[40, 4] = S130;
                oSheet.Cells[41, 4] = S131;
                oSheet.Cells[42, 4] = S132;
                oSheet.Cells[43, 4] = S133;
                oSheet.Cells[44, 4] = S134;
                oSheet.Cells[45, 4] = S135;
                oSheet.Cells[46, 4] = S136;
                oSheet.Cells[47, 4] = S137;
                oSheet.Cells[48, 4] = S138;
                oSheet.Cells[49, 4] = S139;

                oSheet.Cells[50, 4] = S140;
                oSheet.Cells[51, 4] = S141;
                oSheet.Cells[52, 4] = S142;
                oSheet.Cells[53, 4] = S143;
                oSheet.Cells[54, 4] = S144;
                oSheet.Cells[55, 4] = S145;
                oSheet.Cells[56, 4] = S146;
                oSheet.Cells[57, 4] = S147;
                oSheet.Cells[58, 4] = S148;
                oSheet.Cells[59, 4] = S149;

                oSheet.Cells[60, 4] = S150;
                oSheet.Cells[61, 4] = S151;
                oSheet.Cells[62, 4] = S152;
                oSheet.Cells[63, 4] = S153;
                oSheet.Cells[64, 4] = S154;
                oSheet.Cells[65, 4] = S155;
                oSheet.Cells[66, 4] = S156;
                oSheet.Cells[67, 4] = S157;
                oSheet.Cells[68, 4] = S158;
                oSheet.Cells[69, 4] = S159;

                oSheet.Cells[70, 4] = S160;
                oSheet.Cells[71, 4] = S161;
                oSheet.Cells[72, 4] = S162;
                oSheet.Cells[73, 4] = S163;
                oSheet.Cells[74, 4] = S164;
                oSheet.Cells[75, 4] = S165;
                oSheet.Cells[76, 4] = S166;
                oSheet.Cells[77, 4] = S167;
                oSheet.Cells[78, 4] = S168;
                oSheet.Cells[79, 4] = S169;

                oSheet.Cells[80, 4] = S170;
                oSheet.Cells[81, 4] = S171;
                oSheet.Cells[82, 4] = S172;
                oSheet.Cells[83, 4] = S173;
                oSheet.Cells[84, 4] = S174;
                oSheet.Cells[85, 4] = S175;
                oSheet.Cells[86, 4] = S176;
                oSheet.Cells[87, 4] = S177;
                oSheet.Cells[88, 4] = S178;
                oSheet.Cells[89, 4] = S179;

                oSheet.Cells[90, 4] = S180;
                oSheet.Cells[91, 4] = S181;
                oSheet.Cells[92, 4] = S182;
                oSheet.Cells[93, 4] = S183;
                oSheet.Cells[94, 4] = S184;
                oSheet.Cells[95, 4] = S185;
                oSheet.Cells[96, 4] = S186;
                oSheet.Cells[97, 4] = S187;
                oSheet.Cells[98, 4] = S188;
                oSheet.Cells[99, 4] = S189;

                oSheet.Cells[100, 4] = S190;
                oSheet.Cells[101, 4] = S191;
                oSheet.Cells[102, 4] = S192;
                oSheet.Cells[103, 4] = S193;
                oSheet.Cells[104, 4] = S194;
                oSheet.Cells[105, 4] = S195;
                oSheet.Cells[106, 4] = S196;
                oSheet.Cells[107, 4] = S197;
                oSheet.Cells[108, 4] = S198;
                oSheet.Cells[109, 4] = S199;

                // write bach the row counter in Excel
                sReihe = Convert.ToString(iReihe);
                oSheet.Cells[1, 1] = sReihe;
                // and save Excel
                oWB.Save();

            }
            // Quit after saving Excel
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (oWB != null)
                {
                    oWB.Close(true, null, null);
                    oXL.Quit();
                }

            }
        }
    }
}