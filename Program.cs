using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NDbfReader;
using ExcelLibrary.SpreadSheet;

// Install-Package NDbfReader
// https://github.com/eXavera/NDbfReader/blob/master/README.md  -- CHANGED!  It contained ERRORS!!

// Add references:
// ExcelLibrary.dll - https://code.google.com/p/excellibrary/

namespace KESZ_Kalkulator    
{
  class Program
  {
    static DateTime targetDate;                                                        // A halmozás céldátuma

    static void Main(string[] args)
    {
      Console.WriteLine("******************************************************************************");
      Console.WriteLine("***  Szolgáló  ||  Készletérték kalkuláció adott dátumra  ||  (c) eMeL Bt. ***");
      Console.WriteLine("******************************************************************************");
      Console.WriteLine();

      string paramPath = null;
      string paramDate = null;

      if (args.Length > 0)
      {
        Console.WriteLine("Program paraméterek:");

        foreach (string arg in args)
        {
          if ((arg.ToUpper() == "/HELP") || (arg.ToUpper() == "/H") ||
              (arg.ToUpper() == "-HELP") || (arg.ToUpper() == "-H") ||
              (arg.ToUpper() == "--HELP") || (arg.ToUpper() == "--H") ||
              (arg == "/?") || (arg == "-?") || (arg == "--?"))
          {
            Console.WriteLine("A program használata:");
            Console.WriteLine();
            Console.WriteLine("KESZ_Kalkulator.exe [alkönyvtárnév] [cél dátum]");
            Console.WriteLine();
            Console.WriteLine("Az alkönyvtárba kell elhelyezni a KESZ.dbf MOZG.dbf CIKK.dbf és RAKT.dbf állományokat,");
            Console.WriteLine("és ebben a könyvtárban keletkeznek az eredmény XLS állományok.");
            Console.WriteLine("Ha nem határozzuk meg, akkor a .\\ (aktuális) alkönyvtárat használja a program.");
            Console.WriteLine("Ha nem határozunk meg cél dátumot, akkor az bekérésre kerül.");
            Console.WriteLine("A dátum formátuma: ÉÉÉÉ.HH.NN");
            Console.WriteLine();
            Console.WriteLine("Részletesebb információk a www.emel.hu/programok/help/kesz_kalkulator webcímen.");
            Console.ReadLine();
            return;
          }

          //

          string info = "?????  Nem értelmezhető paraméter!";

          Uri uri = new Uri(arg, UriKind.RelativeOrAbsolute);
          DateTime tempDate;

          if ((paramPath == null) && Directory.Exists(arg))
          {
            paramPath = arg;
            info = "létező könyvtárnév";
          }
          else if ((paramPath == null) && uri.IsWellFormedOriginalString())
          {
            paramPath = arg;
            info = "megfelelő könyvtárnév";
          }
          else if ((paramDate == null) && DateTime.TryParse(arg, out tempDate))
          {
            paramDate = arg;
            info = "cél dátum";
          }

          Console.WriteLine("paraméter: '{0}': {1}", arg, info);
        }

        Console.WriteLine();
      }

      //---------------------------------------------------

      if (String.IsNullOrWhiteSpace(paramPath))
      {
        paramPath = @".\";
      }

      if (!Directory.Exists(paramPath))
      {
        Console.WriteLine("A '{0}' alkönyvtár nem létezik!", paramPath);
      }

      string filenameKESZ = Path.Combine(paramPath, "KESZ.DBF");
      string filenameMOZG = Path.Combine(paramPath, "MOZG.DBF");
      string filenameCIKK = Path.Combine(paramPath, "CIKK.DBF");
      string filenameRAKT = Path.Combine(paramPath, "RAKT.DBF");  

      if (!File.Exists(filenameKESZ))
      {
        Console.WriteLine("A '{0}' adatállomány nem létezik!", filenameKESZ);
      }

      if (!File.Exists(filenameMOZG))
      {
        Console.WriteLine("A '{0}' adatállomány nem létezik!", filenameMOZG);
      }

      if (!File.Exists(filenameCIKK))
      {
        Console.WriteLine("A '{0}' adatállomány nem létezik!", filenameCIKK);
      }

      if (!File.Exists(filenameRAKT))
      {
        Console.WriteLine("A '{0}' adatállomány nem létezik!", filenameRAKT);
      }

      //---------------------------------------------------
      
      if (! DateTime.TryParse(paramDate, out targetDate))
      {
        Console.WriteLine("A megadott cél dátum [{0}] érvénytelen!", paramDate);
        Console.WriteLine("Jelenleg a dátumot kötelező parancssorban megadni, még nincs bekérés megvalósítva.");
        return;
      }

      string filenameRESULT = Path.Combine(paramPath, "KESZ.Kalkulalt_" + targetDate.ToString("yyyyMMdd") + ".XLS");           // Eredmény állomány
      
      //---------------------------------------------------

      List<KeszRec> keszRecs = new List<KeszRec>();

      using (var keszTable = Table.Open(filenameKESZ))
      {
        Console.WriteLine("KESZ tábla beolvasása...     [{0} rekord]", keszTable.recCount);

        Reader keszReader = keszTable.OpenReader(Encoding.GetEncoding(852));

        while (keszReader.Read())
        {
          string rsz  = keszReader.GetString("RSZ");
          string csz  = keszReader.GetString("CSZ");

          decimal? temp;
        
          temp = keszReader.GetDecimal("BAR");
          decimal bar  = (temp == null) ? (decimal)0.0 : (decimal)temp;

          temp = keszReader.GetDecimal("NMEN");
          decimal nmen = (temp == null) ? (decimal)0.0 : (decimal)temp;


          KeszRec keszRec = (from kr in keszRecs
                             where (kr.RSZ == rsz) && (kr.CSZ == csz) && (kr.BAR == bar)
                             select kr).FirstOrDefault();

          if (keszRec == null)
          {
            keszRec = new KeszRec(keszReader);

            keszRecs.Add(keszRec);
          }
          else
          {
            keszRec.ZMEN += nmen;
          }
        }
      }

      //---------------------------------------------------

      using (StreamWriter mozgStreamWriter = new StreamWriter(Path.Combine(paramPath, "__KESZ_Kalkulator.MOZG.naplo.txt")),
                          keszStreamWriter = new StreamWriter(Path.Combine(paramPath, "__KESZ_Kalkulator.KESZ.naplo.txt")))
      {
        using (var mozgTable = Table.Open(filenameMOZG))
        {
          Reader mozgReader = mozgTable.OpenReader(Encoding.GetEncoding(852));

          Console.WriteLine("MOZG tábla feldolgozása...   [{0} rekord]", mozgTable.recCount);

          int     datumonTul    = 0;
          decimal kimaradtErtek = 0;

          while (mozgReader.Read())
          {
            MozgRec mozgRec = new MozgRec(mozgReader);

            if (! mozgRec.ZART)
            {
              bool hasznalni = (mozgRec.BDAT <= targetDate);

              mozgRec.WriteNaplo(mozgStreamWriter, hasznalni);

              if (hasznalni)
              {
                KeszRec foundRec = (from keszRec in keszRecs
                                    where (keszRec.RSZ == mozgRec.RSZ) &&
                                          (keszRec.CSZ == mozgRec.CSZ) &&
                                          (keszRec.BAR == mozgRec.BAR)
                                    select keszRec).FirstOrDefault();

                if (foundRec == null)
                {
                  foundRec = new KeszRec(mozgRec, keszStreamWriter);
                  keszRecs.Add(foundRec);
                }
                else
                {
                  foundRec.ZMEN += (mozgRec.MMEN * mozgRec.SIGN);
                }
              }
              else
              {
                datumonTul++;
                kimaradtErtek += (mozgRec.MMEN * mozgRec.SIGN * mozgRec.BAR);
              }
            }
          }

          Console.WriteLine("   Dátumon túli tétel: {0} db",      datumonTul);
          Console.WriteLine("   Dátumon túli érték: {0:0.00} Ft", kimaradtErtek);
        }
      }

      keszRecs.Sort();  

      //---------------------------------------------------

      using (var cikkTable = Table.Open(filenameCIKK))
      {
        Reader cikkReader = cikkTable.OpenReader(Encoding.GetEncoding(852));

        Console.WriteLine("CIKK tábla feldolgozása...   [{0} rekord]", cikkTable.recCount);

        while (cikkReader.Read())
        {
          string csz = cikkReader.GetString("CSZ");

          var foundRecs = from keszRec in keszRecs
                          where (keszRec.CSZ == csz)
                          select keszRec;

          if (foundRecs != null)
          {
            string cmn = cikkReader.GetString("CMN");
            string cme = cikkReader.GetString("CME");

            foreach (var keszRec in foundRecs)
            {
              keszRec.CMN = cmn;
              keszRec.CME = cme;
            }
          }
        }
      }

      //---------------------------------------------------

      List<RaktSum> raktSums = 
                    (from keszRec in keszRecs
                     group keszRec by keszRec.RSZ into rakt
                     orderby rakt.Key
                     select new RaktSum() { RSZ = rakt.Key, ZMEN = rakt.Sum(r => r.ZMEN), ZERT = rakt.Sum(r => r.ZMEN * r.BAR) }).ToList();

      raktSums.Sort();

      //---------------------------------------------------

      using (var raktTable = Table.Open(filenameRAKT))
      {
        Reader raktReader = raktTable.OpenReader(Encoding.GetEncoding(852));

        Console.WriteLine("RAKT tábla feldolgozása...   [{0} rekord]", raktTable.recCount);

        while (raktReader.Read())
        {
          string rsz = raktReader.GetString("RSZ");

          RaktSum foundRakt = (from raktRec in raktSums
                               where (raktRec.RSZ == rsz)
                               select raktRec).FirstOrDefault();

          if (foundRakt != null)
          {
            foundRakt.RMN = raktReader.GetString("RMN");
          }
        }
      }

      //---------------------------------------------------

      Console.WriteLine();
      Console.WriteLine("Eredmény kiiratása Excel állományba...");

      SaveToExcel(filenameRESULT, keszRecs, raktSums);

      //---------------------------------------------------

      Console.WriteLine();
      Console.WriteLine("Program befejezése: <Enter> billentyű leütése.");
      Console.ReadLine();
    }

    //----------------------------------------------------------------------------------------------

    private static void SaveToExcel(string filename, List<KeszRec> keszRecs, List<RaktSum> raktSums)
    {
      Workbook workbook = new Workbook();

      Worksheet worksheet = new Worksheet("Raktárankénti összesen");

      worksheet.Cells[0, 0] = new Cell("Raktár");
      worksheet.Cells[0, 1] = new Cell("Raktár megnevezése");
      worksheet.Cells[0, 2] = new Cell("Záró mennyiség");
      worksheet.Cells[0, 3] = new Cell("Záró érték");

      int rowNo = 1;

      foreach (var item in raktSums)
      {
        worksheet.Cells[rowNo, 0] = new Cell(item.RSZ);
        worksheet.Cells[rowNo, 1] = new Cell(item.RMN);
        worksheet.Cells[rowNo, 2] = new Cell(item.ZMEN);
        worksheet.Cells[rowNo, 3] = new Cell(item.ZERT);

        rowNo++;
      }

      rowNo++;

      var sumZMEN = raktSums.Sum(rs => rs.ZMEN);
      var sumZERT = raktSums.Sum(rs => rs.ZERT);

      worksheet.Cells[rowNo, 0] = new Cell("  ");
      worksheet.Cells[rowNo, 1] = new Cell("Összesen:");
      worksheet.Cells[rowNo, 2] = new Cell(sumZMEN);
      worksheet.Cells[rowNo, 3] = new Cell(sumZERT);

      workbook.Worksheets.Add(worksheet);

      //

      worksheet = new Worksheet("Készlettételek");

      worksheet.Cells[0, 0] = new Cell("Raktár");
      worksheet.Cells[0, 1] = new Cell("Cikkszám");
      worksheet.Cells[0, 2] = new Cell("Cikk megnevezés");
      worksheet.Cells[0, 3] = new Cell("Beszerzési ár");
      worksheet.Cells[0, 4] = new Cell("Záró mennyiség");
      worksheet.Cells[0, 5] = new Cell("Me.");
      worksheet.Cells[0, 6] = new Cell("Záró érték");
      worksheet.Cells[0,12] = new Cell("*");


      rowNo = 1;

      foreach (var item in keszRecs)
      {
        worksheet.Cells[rowNo, 0] = new Cell(item.RSZ);
        worksheet.Cells[rowNo, 1] = new Cell(item.CSZ);
        worksheet.Cells[rowNo, 2] = new Cell(item.CMN);
        worksheet.Cells[rowNo, 3] = new Cell(item.BAR);
        worksheet.Cells[rowNo, 4] = new Cell(item.ZMEN);
        worksheet.Cells[rowNo, 5] = new Cell(item.CME);        
        worksheet.Cells[rowNo, 6] = new Cell(item.ZMEN * item.BAR);
        worksheet.Cells[rowNo,12] = new Cell(item.Eredeti ? "" : "*");

        rowNo++;
      }

      rowNo++;

      sumZERT = keszRecs.Sum(rs => rs.ZMEN * rs.BAR);

      worksheet.Cells[rowNo, 0] = new Cell("");
      worksheet.Cells[rowNo, 1] = new Cell("");
      worksheet.Cells[rowNo, 2] = new Cell("Összesen:");
      worksheet.Cells[rowNo, 3] = new Cell("");
      worksheet.Cells[rowNo, 4] = new Cell("");
      worksheet.Cells[rowNo, 5] = new Cell("");
      worksheet.Cells[rowNo, 6] = new Cell(sumZERT);

      workbook.Worksheets.Add(worksheet);

      //

      List<KeszRec> keszRecsKO =
        (from keszRec in keszRecs
         group keszRec by keszRec.RSZ + keszRec.CSZ into ko
         orderby ko.Key
         select new KeszRec()
         {
           RSZ = ko.First().RSZ,
           CSZ = ko.First().CSZ,
           CMN = ko.First().CMN,
           CME = ko.First().CME,
           ZMEN = ko.Sum(r => r.ZMEN),
           ZERT = ko.Sum(r => r.ZMEN * r.BAR)
         }).ToList();

      keszRecsKO.Sort();

      worksheet = new Worksheet("Készlet összesenek.");

      worksheet.Cells[0, 0] = new Cell("Raktár");
      worksheet.Cells[0, 1] = new Cell("Cikkszám");
      worksheet.Cells[0, 2] = new Cell("Cikk megnevezés");
      worksheet.Cells[0, 3] = new Cell("Záró mennyiség");
      worksheet.Cells[0, 4] = new Cell("Me.");
      worksheet.Cells[0, 5] = new Cell("Záró érték");

      rowNo = 1;

      foreach (var item in keszRecsKO)
      {
        worksheet.Cells[rowNo, 0] = new Cell(item.RSZ);
        worksheet.Cells[rowNo, 1] = new Cell(item.CSZ);
        worksheet.Cells[rowNo, 2] = new Cell(item.CMN);
        worksheet.Cells[rowNo, 3] = new Cell(item.ZMEN);
        worksheet.Cells[rowNo, 4] = new Cell(item.CME);
        worksheet.Cells[rowNo, 5] = new Cell(item.ZERT);

        rowNo++;
      }

      rowNo++;

      sumZERT = keszRecsKO.Sum(rs => rs.ZERT);

      worksheet.Cells[rowNo, 0] = new Cell("");
      worksheet.Cells[rowNo, 1] = new Cell("");
      worksheet.Cells[rowNo, 2] = new Cell("Összesen:");
      worksheet.Cells[rowNo, 3] = new Cell("");
      worksheet.Cells[rowNo, 4] = new Cell("");
      worksheet.Cells[rowNo, 5] = new Cell(sumZERT);

      workbook.Worksheets.Add(worksheet);

      //

      worksheet = new Worksheet("Készült");

      worksheet.Cells[0, 0] = new Cell("Szolgáló Rendszer || Készletérték kalkuláció adott dátumra || eMeL Bt.");

      worksheet.Cells[2, 0] = new Cell("Készítette:");
      worksheet.Cells[2, 1] = new Cell(Environment.UserName);
      worksheet.Cells[2, 2] = new Cell(Environment.UserDomainName);
      worksheet.Cells[2, 3] = new Cell(Environment.MachineName);

      worksheet.Cells[3, 0] = new Cell("Mikor:");
      worksheet.Cells[3, 1] = new Cell(DateTime.Now.ToString());

      worksheet.Cells[5, 0] = new Cell("Készlet dátum:");
      worksheet.Cells[5, 1] = new Cell(targetDate.ToString("d"));
      
      workbook.Worksheets.Add(worksheet);

      //

      workbook.Save(filename);
    }

    //----------------------------------------------------------------------------------------------

    public class KeszRec : IComparable, IComparable<KeszRec>
    {
      public string   RSZ;
      public string   CSZ;
      public string   CMN;
      public string   CME;
      public decimal  BAR;
      public decimal  ZMEN;
      public decimal  ZERT;                                                 // only for LinQ.groupby.sum
      public bool     Eredeti;

      public KeszRec()
      {
        // Empty: for LINQ
      }

      public KeszRec(Reader reader)  
      {
        RSZ  = reader.GetString("RSZ");
        CSZ  = reader.GetString("CSZ");

        decimal? temp;
        
        temp = reader.GetDecimal("BAR");
        BAR  = (temp == null) ? (decimal)0.0 : (decimal)temp;

        temp = reader.GetDecimal("NMEN");
        ZMEN = (temp == null) ? (decimal)0.0 : (decimal)temp;                   // induló mennyiség - ráhalmoz így lesz záró mennyiség       
        
        Eredeti = true;               
      }

      public KeszRec(MozgRec mozgRec, StreamWriter naplo)
      {
        this.RSZ  = mozgRec.RSZ;
        this.CSZ  = mozgRec.CSZ;
        this.BAR  = mozgRec.BAR;

        Eredeti   = false;    

        WriteNaplo(naplo);
      }

      private void WriteNaplo(StreamWriter naplo)
      {
        naplo.WriteLine("Készlettétel létrehozása: {0,-2} {1,-20} {2,20}", RSZ, CSZ, BAR.ToString("0.00"));
      }

      #region IComparable Members

      public int CompareTo(object obj)
      {
        if (obj == null)
        {
          return 1;
        }

        KeszRec kr = obj as KeszRec;
        if (kr != null)
        {
          return CompareTo(kr);
        }
        else
        {
          throw new ArgumentException("Object is not a RaktSum");
        }
      }

      public int CompareTo(KeszRec other)
      { // Key: RSZ+CSZ+BAR+BDAT
        int comp = this.RSZ.CompareTo(other.RSZ);

        if (comp == 0)
        {
          comp = this.CSZ.CompareTo(other.CSZ);

          if (comp == 0)
          {
            comp = this.BAR.CompareTo(other.BAR);
          }
        }
        
        return comp;
      }

      #endregion
    }

    //----------------------------------------------------------------------------------------------

    public class MozgRec
    {
      public string   RSZ;
      public string   CSZ;
      public DateTime BDAT;                                         // Bizonylat(!) dátuma
      public decimal  BAR;
      public decimal  MMEN;    
      public int      SIGN;                                         // -1: készletnövelő mozgástétel; 1:készletcsökkentő mozgásnem
      public bool     ZART;

      public MozgRec(Reader reader)
      { 
        RSZ = reader.GetString("RSZ");
        CSZ = reader.GetString("CSZ");

        BDAT = reader.GetDate("BDAT") ?? DateTime.MinValue;

        {
          decimal? temp;

          temp = reader.GetDecimal("BAR");
          BAR = (temp == null) ? (decimal)0.0 : (decimal)temp;

          temp = reader.GetDecimal("MMEN");
          MMEN = (temp == null) ? (decimal)0.0 : (decimal)temp;
        }

        {
          decimal? temp;

          temp = reader.GetDecimal("SIGN");
          SIGN = (temp == null) ? 1 : ((temp < 0) ? -1 : 1);
        }

        {
          bool? temp;

          temp = reader.GetBoolean("ZART");
          ZART = (temp == null) ? false : (bool)temp;
        }
      }

      public void WriteNaplo(StreamWriter naplo, bool hasznalt)
      {
        string message = hasznalt ? "halmozott." : "dátumon túl, nem halmozott!";

        naplo.WriteLine("Mozgástétel : {0,-2} {1,-20} {2,16} {3,-12} : {4}", RSZ, CSZ, BAR.ToString("0.00"), BDAT.ToString("d"), message);
      }
    }

    //----------------------------------------------------------------------------------------------

    public class RaktSum : IComparable, IComparable<RaktSum>
    {
      public string  RSZ;
      public string  RMN;
      public decimal ZMEN;
      public decimal ZERT;

      #region IComparable Members

      public int CompareTo(object obj)
      {
        if (obj == null)
        {
          return 1;
        }

        RaktSum rs = obj as RaktSum;
        if (rs != null)
        {
          return CompareTo(rs);
        }
        else
        {
          throw new ArgumentException("Object is not a RaktSum");
        }
      }

      public int CompareTo(RaktSum other)
      {
        return this.RSZ.CompareTo(other.RSZ);
      }

      #endregion
    }
  }
}
