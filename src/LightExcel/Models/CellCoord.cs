using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LightExcel.Models;

public readonly record struct CellCoord(int? X, int? Y);