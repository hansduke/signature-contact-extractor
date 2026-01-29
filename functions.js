function extractAndCreateContact() {
    Office.onReady((info) => {
        if (info.host === Office.HostType.Outlook) {
            Office.addin.showAsTaskpane();
        }
    });
    Office.context.ui.closeContainer();
}
