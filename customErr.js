/*
    FrM Futech 2025-03-18
    Codes d'erreur personnalisés
*/

function customException(code, message){
    this.message = message;
    this.code = code;
    this.type = "Exception personnalisée";
}