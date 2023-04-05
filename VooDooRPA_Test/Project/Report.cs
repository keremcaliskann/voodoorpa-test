using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VooDooRPA_Project
{
    class Report
    {
        public int siraNo;
        public int adet;
        public float kgDesi;
        public string mesafe;
        public float ucret;

        public Report(int _siraNo, int _adet, float _kgDesi, string _mesafe, float _ucret)
        {
            siraNo = _siraNo;
            adet = _adet;
            kgDesi = _kgDesi;
            mesafe = _mesafe;
            ucret = _ucret;
        }
    }
}
