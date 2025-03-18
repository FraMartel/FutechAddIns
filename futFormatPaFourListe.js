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
      /** Récupérer la référence à l'onglet*/
      let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
      
      /** Vérifier que le document est valide */
      // Nom de l'onglet
      selectedSheet.load("name");
      await context.sync();
      if(selectedSheet.name != 'FactureAPayerTable'){
        throw new customException(5000, "Nom de feuille invalide");
      };
    
      // Ordre et titres de l'entête
      let rEnteteOriginal = selectedSheet.getRange("A5:K5");
      let rEnteteModif = selectedSheet.getRange("A1:K1");
      rEnteteOriginal.load("text");
      rEnteteModif.load("text");
      await context.sync();
      if(!(checkEntete(rEnteteOriginal,1) && checkEntete(rEnteteModif,0))){
        throw new customException(5001, "Entêtes absents ou dans le mauvais ordre, fichier incompatible.");
      };
    
      /** Ajuster la mise en page de la feuille */
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

      //** Fonctions utilitaires */
      function checkEntete(rEntete, isOriginal){
        // Vérifiction des entêtes, originale ou modifiée, pour valider l'origine du fichier
        let errCount = 0;
        return ((isOriginal = 1 && rEntete[0] == "<input type='checkbox' >") || (isOriginal = 0 && rEntete[0] == ""))
              && rEntete[1] == "N° facture"
              && rEntete[2] == "N° transaction"
              && rEntete[3] == "Référence"
              && rEntete[4] == "Montant"
              && rEntete[5] == "Paiement"
              && rEntete[6] == "Esc. $"
              && rEntete[7] == "Solde"
              && rEntete[8] == "Date"
              && rEntete[9] == "Échéance"
              && ((isOriginal = 1 && rEntete[10] == "Terme paiement") || (isOriginal = 0 && rEntete[10] == "Terme"))
      };

    });
    
  } catch (error) {
      //Gérer les erreurs ici...
      console.error(error);
  };

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
};


  
  // Register the function with Office.
  Office.actions.associate("futFormatPaFourListe", futFormatPaFourListe);
  