using ShapeCrawler;

namespace SlideAssembler.Operations
{
    public class SetWidth : IPrestationOperation
    {
        private readonly string name;
        private decimal data;

        public SetWidth(string name, decimal data)
        {
            this.name = name;
            this.data = data;
        }
        public void Apply(Presentation presentation)
        {
            foreach (var slide in presentation.Slides)
            {
                foreach (var shape in slide.Shapes)
                {
                    if (shape.Name.Equals(name))
                    {
                        shape.Width = data;
                        shape.TextFrame.Text = data.ToString();
                    }
                }
            }
        }
    }
}
