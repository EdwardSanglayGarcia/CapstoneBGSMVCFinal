using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MVCCapstoneBGS
{
    public class DefaultData
    {
        public string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/Capstone Generated Document";

        public string CONSUMER_KEY = "N5rwRYcClshCYS0qW0qsySmvx";
        public string CONSUMER_SECRET = "57ywSvpsf0jLvdYQIaJLQJYveufqDI4QqmFoGLcl56UbeRHsvq";
        public string ACCESS_TOKEN = "1258097237348896769-AL5rIiYea1qvhRS2PDzUZvrc0CoNAz";
        public string ACCESS_TOKEN_SECRET = "eg93G03bbzry7BrFT29Qqe6f8mNe7dbnJ4lvD3cOmKAuk";

        public string MAP_A = @"https://api.tiles.mapbox.com/v4/mapbox.streets/{z}/{x}/{y}.png?access_token=pk.eyJ1IjoiYmJyb29rMTU0IiwiYSI6ImNpcXN3dnJrdDAwMGNmd250bjhvZXpnbWsifQ.Nf9Zkfchos577IanoKMoYQ";
        public string MAP_B = @"https://api.mapbox.com/styles/v1/mapbox/streets-v11/tiles/{z}/{x}/{y}?access_token=pk.eyJ1IjoibWFwYm94IiwiYSI6ImNpejY4NXVycTA2emYycXBndHRqcmZ3N3gifQ.rJcFIG214AriISLbB6B5aw";

    }
}