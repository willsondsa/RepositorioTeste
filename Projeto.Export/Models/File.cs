using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Projeto.Export.Models
{
    public class UpFile
    {
      public HttpPostedFile Posted { get; set; }
      public DateTime DateModify { get; set; }
    }
}