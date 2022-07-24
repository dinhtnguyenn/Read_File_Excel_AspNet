using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebApplication1.Models
{
    public class Info
    {
        public Info(string stt, string name, string classes)
        {
            this.stt = stt;
            this.name = name;
            this.classes = classes;
        }

        public string stt { get; set; }
        public string name { get; set; }
        public string classes { get; set; }


    }
}