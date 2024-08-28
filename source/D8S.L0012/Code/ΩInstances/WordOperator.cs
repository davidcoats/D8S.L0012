using System;


namespace D8S.L0012
{
    public class WordOperator : IWordOperator
    {
        #region Infrastructure

        public static IWordOperator Instance { get; } = new WordOperator();


        private WordOperator()
        {
        }

        #endregion
    }
}
