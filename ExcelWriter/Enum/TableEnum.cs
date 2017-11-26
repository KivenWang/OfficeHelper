﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelWriterCSharp
{
    public enum TableStyle
    {
        None,
        Light1,
        Light2,
        Light3,
        Light4,
        Light5,
        Light6,
        Light7,
        Light8,
        Light9,
        Light10,
        Light11,
        Light12,
        Light13,
        Light14,
        Light15,
        Light16,
        Light17,
        Light18,
        Light19,
        Light20,
        Light21,
        Medium1,
        Medium2,
        Medium3,
        Medium4,
        Medium5,
        Medium6,
        Medium7,
        Medium8,
        Medium9,
        Medium10,
        Medium11,
        Medium12,
        Medium13,
        Medium14,
        Medium15,
        Medium16,
        Medium17,
        Medium18,
        Medium19,
        Medium20,
        Medium21,
        Medium22,
        Medium23,
        Medium24,
        Medium25,
        Medium26,
        Medium27,
        Medium28,
        Dark1,
        Dark2,
        Dark3,
        Dark4,
        Dark5,
        Dark6,
        Dark7,
        Dark8,
        Dark9,
        Dark10,
        Dark11
    }

    public enum TableHeaders
    {
        Yes = Excel.XlYesNoGuess.xlYes,
        No = Excel.XlYesNoGuess.xlNo,
    }
}
