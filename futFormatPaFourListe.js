/*
    FrM Futech 2025-03-04
    Fonction futFormatPaFourListe
      Pour complément Excel
    Formate la liste des paiements fournisseur pour contre-vérification des factures fournisseur.
    Formaté pour impression selon les standards établis par NaG/AnL, mars 2025
*/

Office.onReady(() => {
    
});
  
async function futFormatPaFourListe(event) {
  try {
    await Excel.run(async (context) => {
    /** Ajuster la mise en page de la feuille */
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

    // Set print area for selectedSheet to range "A:K"
    selectedSheet.pageLayout.setPrintArea("A:K");
    // Set ExcelScript.PageOrientation.landscape orientation for selectedSheet
    selectedSheet.pageLayout.orientation = Excel.PageOrientation.landscape;
    // Répéter seulement la rangée 5 sur toutes les pages
    selectedSheet.pageLayout.setPrintTitleRows("$5:$5");
    // Set Letter paperSize for selectedSheet
    selectedSheet.pageLayout.paperSize = Excel.PaperType["letter"];
    // Set FitAllColumnsOnOnePage scaling for selectedSheet
    selectedSheet.pageLayout.zoom = { horizontalFitToPages: 1, verticalFitToPages: 0, scale: null };

    await context.sync();
    });
  } catch (error) {
      //Gérer les erreurs ici...
      console.error(error);
  }

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}
  
  // Register the function with Office.
  Office.actions.associate("futFormatPaFourListe", futFormatPaFourListe);
  