using System;
using System.Threading.Tasks;

using R5T.T0132;

using D8S.L0011;


namespace D8S.L0012
{
    [FunctionalityMarker]
    public partial interface IWordOperator : IFunctionalityMarker
    {
        public void In_ApplicationContext_Synchronous(
            Action<Application> applicationAction)
        {
            using var application = new Application();

            applicationAction(application);
        }

        public async Task In_ApplicationContext_Asynchronous(
            Func<Application, Task> applicationAction)
        {
            using var application = new Application();

            await applicationAction(application);
        }

        /// <summary>
        /// Chooses <see cref="In_ApplicationContext_Asynchronous(Func{Application, Task})"/> as the default.
        /// </summary>
        public Task In_ApplicationContext(
            Func<Application, Task> applicationAction)
            => this.In_ApplicationContext_Asynchronous(
                applicationAction);


        public async Task In_DocumentContext_New(
            Func<Document, Application, Task> documentAction)
        {
            await this.In_ApplicationContext(
                async application =>
                {
                    var document = application.New_Document();

                    await documentAction(
                        document,
                        application);
                });
        }

        public async Task In_DocumentContext_New(
            string wordDocumentFilePath,
            Func<Document, Application, Task> documentAction)
        {
            async Task Internal(Document document, Application application)
            {
                await documentAction(document, application);

                document.SaveAs(wordDocumentFilePath);
            }

            await this.In_DocumentContext_New(
                Internal);
        }

        public async Task In_DocumentContext_New(
            Func<Document, Task> documentAction)
        {
            await this.In_ApplicationContext(
                async application =>
                {
                    var document = application.New_Document();

                    await documentAction(
                        document);
                });
        }

        public async Task In_DocumentContext_New(
            string wordDocumentFilePath,
            Func<Document, Task> documentAction)
        {
            async Task Internal(Document document)
            {
                await documentAction(document);

                document.SaveAs(wordDocumentFilePath);
            }

            await this.In_DocumentContext_New(
                Internal);
        }


        public void In_DocumentContext_New(
            Action<Document, Application> documentAction)
        {
            this.In_ApplicationContext_Synchronous(
                application =>
                {
                    var document = application.New_Document();

                    documentAction(
                        document,
                        application);
                });
        }

        public void In_DocumentContext_New(
            string wordDocumentFilePath,
            Action<Document, Application> documentAction)
        {
            void Internal(Document document, Application application)
            {
                documentAction(document, application);

                document.SaveAs(wordDocumentFilePath);
            }

            this.In_DocumentContext_New(
                Internal);
        }

        public void In_DocumentContext_New(
            Action<Document> documentAction)
        {
            this.In_ApplicationContext_Synchronous(
                application =>
                {
                    var document = application.New_Document();

                    documentAction(
                        document);
                });
        }

        public void In_DocumentContext_New(
            string wordDocumentFilePath,
            Action<Document> documentAction)
        {
            void Internal(Document document)
            {
                documentAction(document);

                document.SaveAs(wordDocumentFilePath);
            }

            this.In_DocumentContext_New(
                Internal);
        }

        public void Open_Document_ForInspection(
            string wordExecutableFilePath,
            string wordDocumentFilePath)
        {
            var argumentsString_Enquoted = Instances.StringOperator.Ensure_Enquoted(
                wordDocumentFilePath);

            Instances.ProcessOperator.Start(
                wordExecutableFilePath,
                argumentsString_Enquoted);
        }
    }
}
