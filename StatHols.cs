        /// <summary>
        /// Modifiers:
        ///  DOW = Equivalent day of week, 12-Jan-1900 would always be on 2nd Thu of the month
        ///  EASTER = Generates date of Easter Friday for the year
        ///  FWD = On the date supplied, or next working day
        ///  NWD = Next Working Day
        /// </summary>

        public class StatHoliday : IComparable<StatHoliday>
        {
            public enum Modifiers { DOW, EASTER, FWD, NWD }
 
            private string _Description;
            private DateTime _BaseDate;
            private Modifiers _Modifier;
            private int _YearOffset;
            private string _DateFormat = "yyyy-MMM-dd";
            public string Description { get => _Description; set => _Description = value; }
            public string DateFormat { get => _DateFormat; set => _DateFormat = value; }
 
            public StatHoliday(string Description, DateTime BaseDate, Modifiers Modifier, int YearOffset = 0)
            {
                this.Description = Description;
                _BaseDate = BaseDate;
                _Modifier = Modifier;
                _YearOffset = YearOffset;
            }
 
            public DateTime ThisYear => CorrectedDate(DateTime.Today.Year + _YearOffset);
 
            public DateTime NextYear => CorrectedDate(DateTime.Today.Year + _YearOffset + 1);
 
            public DateTime NextDueDate
            {
                get
                {
                    DateTime chkDate = CorrectedDate(DateTime.Today.Year + _YearOffset);
                    return (chkDate >= DateTime.Today ? chkDate : CorrectedDate(DateTime.Today.Year + _YearOffset + 1));
                }
            }
 
            public DateTime CorrectedDate(int setYear)
            {
                DateTime rc = _BaseDate;
                rc = rc.AddYears(setYear - rc.Year);
                int dow = (int)rc.DayOfWeek;
                if (--dow < 0) dow = 6;
 
                switch (_Modifier)
                {
                    case Modifiers.DOW:
                        int wkDy = (int)_BaseDate.DayOfWeek;
                        int wkNo = (int)(_BaseDate.Day / 7); // + 1;
                        rc = rc.AddDays(1 - rc.Day);
                        int baseWkDy = (int)rc.DayOfWeek;
                        rc = rc.AddDays(wkDy - baseWkDy + (baseWkDy > wkDy ? 7 : 0) + wkNo * 7);
                        break;
                    case Modifiers.EASTER:
                        // excel formula:
                        // =DATE($A$1,3,26+
                        //   (
                        //       MOD((203+($A$1>1899))-MOD($A$1,19)*11,30)                    // bit1
                        //     -(MOD((203+($A$1>1899))-MOD($A$1,19)*11,30)>27)                // bit2
                        //   )
                        //  -MOD($A$1+INT($A$1/4)+(MOD((203+($A$1>1899))-MOD($A$1,19)*11,30)
                        //   -(MOD((203+($A$1>1899))-MOD($A$1,19)*11,30)>27))-(12+($A$1>1899)+($A$1>2099)),7))
                        rc = new DateTime(setYear, 3, 26);
                        int bit1 = (203 + (setYear > 1899 ? 1 : 0) - (setYear % 19) * 11) % 30;
                        int bit2 = ((203 + (setYear > 1899 ? 1 : 0) - (setYear % 19) * 11) % 30 > 27 ? 1 : 0);
 
                        int daysInc = 0;
                        daysInc = bit1 - bit2;
                        daysInc -= (setYear + (int)(setYear / 4) + bit1 - bit2 - (12 + (setYear > 1899 ? 1 : 0) + (setYear > 2099 ? 1 : 0))) % 7;
                        rc = rc.AddDays(daysInc);
                        break;
                    case Modifiers.FWD:
                        rc = rc.AddDays(dow > 4 ? 7 - dow : 0);
                        break;
                    case Modifiers.NWD:
                        rc = rc.AddDays(dow > 3 ? 8 - dow : 1);
                        break;
                    default:
 
                        break;
                }
                return rc;
            }
 
            public override string ToString()
            {
                return $"{ThisYear.ToString(DateFormat)} - {Description}";
            }
 
            public int CompareTo(StatHoliday other)
            {
                return this.ThisYear.CompareTo(other.ThisYear);
            }
        }
