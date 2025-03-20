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
      let wsheet = context.workbook.worksheets.getActiveWorksheet();
      
      /** Vérifier que le document est valide */
      // Nom de l'onglet
      wsheet.load("name");
      await context.sync();
      if(wsheet.name != 'FactureAPayerTable'){
        throw new customException(5000, "Nom de feuille invalide");
      };
    
      // Ordre et titres de l'entête
      let rEnteteOriginal = wsheet.getRange("A5:K5");
      let rEnteteModif = wsheet.getRange("A1:K1");
      rEnteteOriginal.load("text");
      rEnteteModif.load("text");
      await context.sync();
      if(!(checkEntete(rEnteteOriginal.text[0],1) || checkEntete(rEnteteModif.text[0],0))){
        throw new customException(5001, "Entêtes absents ou dans le mauvais ordre, fichier incompatible.");
      };
    
      /** Supprimer les rangées et les formats superflus - modifications destructrices */
      // Suppression de l'image (logo Futech)
      let shapes = wsheet.shapes;
      shapes.load("items/$none");
      await context.sync();

      shapes.items.forEach(function (shape) {
        shape.delete();
      });
      
      // Suppression des 4 premières rangées (entête déplacée à rangée 1)
      // seulement si ça n'a pas déjà été fait.
      if(checkEntete(rEnteteOriginal.text[0],1)){
        let rAvantEntete = wsheet.getRange("1:4");
        rAvantEntete.delete(Excel.DeleteShiftDirection.up);
      };
      
      /** Modifier les entêtes et les largeurs de colonnes - non-destructeur */
      wsheet.getRange("A1").values = [[""]];
      wsheet.getRange("K1").values = [["Terme"]];
      wsheet.getRange("A1").format.columnWidth = 2;
      wsheet.getRange("B1:C1").format.columnWidth = 78;
      wsheet.getRange("D1").format.columnWidth = 150;
      wsheet.getRange("E1:J1").format.columnWidth = 78;
      wsheet.getRange("K1").format.columnWidth = 50;
      wsheet.getRange("E1:J1").format.horizontalAlignment = "Right";
      // Appliquer un format de nombre plus facile à lire sur les montants
      wsheet.getRange("E:H").numberFormat = [["# ### ##0.00"]];

      /** Modifier les couleurs et formats des rangées - non destructeur */
      // Création des objets de format, définition plus bas.
      const oFormatStd = FormatStd();
      const oFormatBold = FormatBold();

      let rUsedRange = wsheet.getUsedRange();
      rUsedRange.load("rowCount,values");
      await context.sync();
  
      rUsedRange.values.forEach( (value, index) => {

        let currentRowFormat = rUsedRange.getRow(index).format;
        let currentRowBorders = currentRowFormat.borders;
        if (index == 0){
          currentRowFormat.set(oFormatBold);
          currentRowBorders.getItem(Excel.BorderIndex.edgeBottom)
          .set({
            style: Excel.BorderLineStyle.continuous,
            weight: Excel.BorderWeight.thick
          });
        } else if (typeof (value[0]) == "boolean"){
          // Action pour les lignes de détails
          currentRowFormat.set(oFormatStd);
          currentRowBorders.getItem(Excel.BorderIndex.edgeBottom)
            .set({
              style: Excel.BorderLineStyle.continuous,
              weight: Excel.BorderWeight.thin
            });
        } else if( value[0].slice(0,5) == "Total"){
          if(index < rUsedRange.rowCount -1){
            // Action pour les totaux partiels
            currentRowFormat.set(oFormatBold);
            currentRowBorders.getItem(Excel.BorderIndex.edgeBottom)
            .set({
              style: Excel.BorderLineStyle.double
            });
          } else {
            // Action pour le total final
            currentRowFormat.set(oFormatBold);
            currentRowBorders.getItem(Excel.BorderIndex.edgeBottom)
            .set({
              style: Excel.BorderLineStyle.double,
              weight: Excel.BorderWeight.medium
            });
          };
        } else {
          // Action pour les lignes de code fournisseur (entêtes)
          currentRowFormat.set(oFormatBold);
          currentRowFormat.rowHeight = 28;
          currentRowBorders.getItem(Excel.BorderIndex.edgeBottom)
            .set({
              style: Excel.BorderLineStyle.continuous,
              weight: Excel.BorderWeight.thin
            });
        };
      });


      /** Ajuster la mise en page de la feuille - modifications non destructrices */
      // Set print area for wsheet to range "A:K"
      wsheet.pageLayout.setPrintArea("A:K");
      // Set ExcelScript.PageOrientation.landscape orientation for wsheet
      wsheet.pageLayout.orientation = Excel.PageOrientation.portrait;
      // Répéter seulement la rangée 5 sur toutes les pages
      wsheet.pageLayout.setPrintTitleRows("$5:$5");
      // Set Letter paperSize for wsheet
      wsheet.pageLayout.paperSize = Excel.PaperType["letter"];
      // Set FitAllColumnsOnOnePage scaling for wsheet
      wsheet.pageLayout.zoom = { horizontalFitToPages: 1, verticalFitToPages: 0, scale: null };
      // Spécifier les marges (marges fines)
      wsheet.pageLayout.setPrintMargins("Centimeters", { bottom: 1.9, top: 1.9, left: 0.6, right: 0.6 });

      await context.sync();

      //** Fonctions utilitaires */
      function checkEntete(rEntete, isOriginal){
        // Vérifiction des entêtes, originale ou modifiée, pour valider l'origine du fichier
        return (((isOriginal == 1 && rEntete[0] == "<input type='checkbox' >") || (isOriginal == 0 && rEntete[0] == ""))
              && rEntete[1] == "N° facture"
              && rEntete[2] == "N° transaction"
              && rEntete[3] == "Référence"
              && rEntete[4] == "Montant"
              && rEntete[5] == "Paiement"
              && rEntete[6] == "Esc. $"
              && rEntete[7] == "Solde"
              && rEntete[8] == "Date"
              && rEntete[9] == "Échéance"
              && ((isOriginal == 1 && rEntete[10] == "Terme paiement") || (isOriginal == 0 && rEntete[10] == "Terme")));
      };

      // Définition du format de l'entête
      function FormatStd(){
        return {
          fill: {color: "FFFFFF"},
          font: {bold: false, color: "black",name: "Tahoma"}
        };
      };
      function FormatBold(){
        return {
          fill: {color: "FFFFFF"},
          font: {bold: true, color: "black",name: "Tahoma"}
        };
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
  