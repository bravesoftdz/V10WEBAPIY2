unit UCreCommun;

interface

uses
    HMsgBox,
    Graphics,
    HEnt1,
    Ent1;


const
   // Taille des colonnes
   cInLargColTaux    = 57;
   cInLargColMnt     = 98;
   cInLargColDate    = 74;

   // Couleurs
   cInModifColor     = clRed;
   cInFondExColor    = clBtnFace;


var
   // Devises
   gCodeDevDossier : String;
   gCodeDevEuro    : String;

   gCodeDevFranc   : String;
   gSymbDevEuro    : String;
   gSymbDevFranc   : String;
   gSymbDevNull    : String;

   gCodeDevEmprunt : String;
   gSymbDevEmprunt : String;
   gLibDevEmprunt  : String;

   gCodeDevSaisie  : String;
   gSymbDevSaisie  : String;
   gLibDevSaisie   : String;


Function EstSerialise : Boolean;


implementation


Function EstSerialise : Boolean;
begin
   Result := True;
   Exit;

   Result := VH^.OkModCre;
   if not Result then
      PGIBox('Ce produit n''est pas sérialisé !',TitreHalley);
end;




end.
