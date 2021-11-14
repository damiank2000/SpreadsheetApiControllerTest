namespace ClosedXmlTest
{
    internal class SampleDto
    {
        public SampleDto(string firstColumn, int secondIntColumn, string thirdColumn)
        {
            FirstColumn = firstColumn;
            SecondIntColumn = secondIntColumn;
            ThirdColumn = thirdColumn;
        }
        public string FirstColumn { get; set; }
        public int SecondIntColumn { get; set; }
        public string ThirdColumn { get; set; }
    }
}