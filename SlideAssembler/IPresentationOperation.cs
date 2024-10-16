using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ShapeCrawler;

namespace SlideAssembler
{
    public interface IPresentationOperation
    {
        void Apply(Presentation presentation);
    }
}