unit UTofGCRepriseDonnees;

interface
Uses StdCtrls, 
     Controls,
     Classes,
     M3FP,
{$IFNDEF EAGLCLIENT}
     db, 
     dbtables, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF ,
     UTOB ,
     SaisieParamGPAO;

Type

     TOF_GCRepriseDonnees = Class (TOF)
       private
       public
       procedure OnArgument(stArgument : String ) ; override ;
       procedure OnLoad                   ; override ;

     END ;


implementation
/////////////////////////////////////////////////////////////////////////////
procedure TOF_GCRepriseDonnees.OnArgument (stArgument : String ) ;
BEGIN
  
END;

////////// Paramètrage Recup Auto Tarif Four ///////////////////////////////////
procedure GCRepriseDonnees (parms: array of variant; nb: integer );
begin
if parms [1] = '' then AppelParamGPAO (taCreat, parms [1], parms [2], parms[3])
else AppelParamGPAO (taModif, parms [1], parms [2], parms [3]);
end;

////////// Paramètrage Recup Auto Tarif Four ///////////////////////////////////
procedure GCSuppRepriseDonnees (parms: array of variant; nb: integer );
begin
if parms [1] <> '' then SupprimeParamGPAO (parms [1], parms [2]);
end;


/////////////////////////////////////////////////////////////////////////////

procedure InitTOFGCRepriseDonnees ;
begin
RegisterAglProc ( 'GCRepriseDonnees', TRUE , 1, GCRepriseDonnees);
RegisterAglProc ( 'GCSuppRepriseDonnees', TRUE , 1, GCSuppRepriseDonnees);
end;

//
//
//
procedure TOF_GCRepriseDonnees.OnLoad;
var Q            : TQuery ;
    Tob_Entete   : TOB    ;
    Tob_Choixcod : TOB ;
begin
  inherited;
  //
  // Désactivation des combos
  //
  if TCheckBox(GetControl('SAISIE_INVENTAIRE')).Checked then
  begin
    //
    // Création des valeurs dans CHOIXCOD
    //
    Q := OpenSQL ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="GT1" AND CC_CODE="..."' ,True) ;
    if Q.EOF then
    begin
      Ferme (Q) ;
      Tob_ChoixCod := TOB.CREATE ('CHOIXCOD', NIL, -1);
      Tob_ChoixCod.PutValue ('CC_TYPE'      , 'GT1');
      Tob_ChoixCod.PutValue ('CC_CODE'      , '...');
      Tob_ChoixCod.PutValue ('CC_LIBELLE'   , 'Reprise inventaire');
      Tob_ChoixCod.PutValue ('CC_ABREGE'    , 'Rep. inventaire');
      Tob_ChoixCod.InsertDB (nil, True);
      Tob_ChoixCod.free;
    end else Ferme (Q) ;
    //
    // Création des valeurs dans CHOIXCOD
    //
    Q := OpenSQL ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="GT2" AND CC_CODE="..."' ,True) ;
    if Q.EOF then
    begin
      Ferme (Q) ;
      Tob_ChoixCod := TOB.CREATE ('CHOIXCOD', NIL, -1);
      Tob_ChoixCod.PutValue ('CC_TYPE'      , 'GT2');
      Tob_ChoixCod.PutValue ('CC_CODE'      , '...');
      Tob_ChoixCod.PutValue ('CC_LIBELLE'   , 'Reprise inventaire');
      Tob_ChoixCod.PutValue ('CC_ABREGE'    , 'Rep. inventaire');
      Tob_ChoixCod.InsertDB (nil, True);
      Tob_ChoixCod.free;
    end else Ferme (Q) ;
    //
    // Création de la reprise inventaire si besoin
    //
    Q := OpenSQL ('SELECT * FROM REPRISEGPAO WHERE GRE_TYPEREPRISE="..." AND GRE_TYPEDONNEES="..."' ,True) ;
    if Q.EOF then
    begin
      Ferme (Q) ;
      //
      // Création de la TOB de reprise "entête"
      //
      Tob_Entete := TOB.CREATE ('REPRISEGPAO', NIL, -1);
      Tob_Entete.PutValue ('GRE_TYPEREPRISE' , '...'        );
      Tob_Entete.PutValue ('GRE_TYPEDONNEES' , '...'        );
      Tob_Entete.PutValue ('GRE_NOMTABLE'    , 'TRANSINVLIG');
      Tob_Entete.PutValue ('GRE_DEBUTENREG'  , 'INV'        );
      Tob_Entete.PutValue ('GRE_LONGUEURFIXE', 'X'          );
      Tob_Entete.PutValue ('GRE_SEPARENREG'  , ''           );
      Tob_Entete.PutValue ('GRE_ENTETELIG'   , 'X'          );
      Tob_Entete.InsertDB (nil, True);
      Tob_Entete.free;
    end else Ferme (Q) ;
    //
    // On force les valeurs d'entrée
    //
    SetControlText('GRE_TYPEREPRISE' , '...');
    SetControlText('GRE_TYPEDONNEES' , '...');

    SetControlEnabled('GRE_TYPEREPRISE', False);
    SetControlEnabled('GRE_TYPEDONNEES', False);
    SetControlEnabled('FFILTRES'       , False);
    SetControlEnabled('BDELETE'        , False);
  end ;
end;

Initialization
registerclasses([TOF_GCRepriseDonnees]) ;
InitTOFGCRepriseDonnees;
end.
