{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 26/01/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : EXCEPTPRESENCESAL
                 Saisie des exceptions de présence d'un salarié
Mots clefs ... : TOM;EXCEPTPRESENCESAL
*****************************************************************
PT1  23/07/2007  FLO  Gestion des anomalies
PT2  10/08/2007  FLO  Recalcul automatique des compteurs lors d'une saisie
}
Unit UTOMExceptpresencesal;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche, 
     FichList, 
{$else}
     eFiche, 
     eFichList, 
{$ENDIF}
     forms, 
     ComCtrls,
     UTOM,
     UTob ;

Type
  TOM_EXCEPTPRESENCESAL = Class (TOM)
    procedure OnUpdateRecord             ; override ;
    procedure OnDeleteRecord             ; override ; //PT2
    procedure OnAfterUpdateRecord        ; override ; //PT2
    procedure OnAfterDeleteRecord        ; override ; //PT2
    procedure OnLoadRecord               ; override ;
    procedure OnArgument ( S: String )   ; override ;
  private
    mode,codesal,libelle,prenom, nomsal, datedeb, datefin : string;
    DateModif : TDateTime; //PT2
    procedure TrtTypeaffect(sender: TObject);
  end ;

Implementation

Uses HCtrls,
     HEnt1,
     HMsgBox,
     sysutils,
     UTOFPGEXCEPTPRESSALGRP, PGPresence;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... :   /  /    
Modifié le ... :   /  /    
Description .. : Création ou modification d'un enregistrement
Mots clefs ... : 
*****************************************************************}
procedure TOM_EXCEPTPRESENCESAL.OnUpdateRecord ;
var
     Salarie,TypeAffect,CycleAffect : String;
     DateDebut, DateFin : TDateTime;
     modeInsert : Boolean;
     GestionPres : TGestionPresence;
     ExceptionSal,ExceptionCycle : Boolean;
     Cycle, TypeCycle : String;
begin
  Inherited ;

     //PT1 - Début
     modeInsert  := True;

     Salarie     := GetControlText('PYE_SALARIE');
     DateDebut   := StrToDate(GetControlText('PYE_DATEDEBUT'));
     DateFin     := StrToDate(GetControlText('PYE_DATEFIN'));
     TypeAffect  := GetControlText('PYE_TYPEAFFECT');
     CycleAffect := GetControlText('PYE_CYCLEAFFECT');

     // Tous les champs doivent être renseignés
     If (Salarie = '') Or (DateDebut=iDate1900) Or (DateFin = idate1900) Or (TypeAffect = '') Or (CycleAffect = '') Then
     Begin
          PGIError (TraduireMemoire('Tous les champs doivent être renseignés.'));
          LastError := 1;
          Exit;
     End;

     If DateDebut > DateFin Then
     Begin
          PGIError (TraduireMemoire('La date de début doit être inférieure ou égale à la date de fin.'));
          LastError := 1;
          Exit;
     End;

     // Contrôles de cohérence
     If ds.State = dsEdit Then modeInsert:=False;
     If Not ControleExisteException (Salarie, DateDebut, DateFin, TypeAffect, modeInsert) Then
     Begin
          PGIError (TraduireMemoire('Une exception existe déjà sur cette période pour le salarié.'));
          LastError := 1;
          Exit;
     End;

     GestionPres := TGestionPresence.Create(True,True,True,True,True,DateDebut,DateFin);
     GestionPres.GetJourneeTypeSalarie(DateDebut, Salarie, Cycle, TypeCycle, ExceptionSal, ExceptionCycle, False);
     If (TypeAffect = TypeCycle) And (Cycle = CycleAffect) AND (ExceptionSal = false) Then
     Begin
          PGIError (TraduireMemoire('Le cycle normal du salarié est identique à l''exception saisie.'));
          LastError := 1;
          FreeAndNil(GestionPres);
          Exit;
     End;
     FreeAndNil(GestionPres);

     If ControleExisteAbsence (Salarie, DateDebut) Then
     Begin
          PGIInfo (TraduireMemoire('Le salarié est absent durant la période sélectionnée.'),'Information');
     End;
     //PT1 - Fin

     //PT2 - Début
     If (LastError = 0) Then
     Begin
          CodeSal   := Salarie;
          DateModif := DateDebut;
     End;
     //PT2 - Fin
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 10/08/2007 / PT2
Modifié le ... :   /  /    
Description .. : Suppression d'un enregistrement
Mots clefs ... : 
*****************************************************************}
procedure TOM_EXCEPTPRESENCESAL.OnDeleteRecord ;
begin
  Inherited ;
     If (LastError = 0) Then
     Begin
          CodeSal   := GetField('PYE_SALARIE');
          DateModif := GetField('PYE_DATEDEBUT');
     End;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 10/08/2007 / PT2
Modifié le ... :   /  /
Description .. : Suite à la création ou modification d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_EXCEPTPRESENCESAL.OnAfterUpdateRecord ;
begin
  Inherited ;
     If (LastError = 0) Then CompteursARecalculer(DateModif, CodeSal);
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 10/08/2007 / PT2
Modifié le ... :   /  /
Description .. : Suite à la suppression d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_EXCEPTPRESENCESAL.OnAfterDeleteRecord ;
begin
  Inherited ;
     If (LastError = 0) Then CompteursARecalculer(DateModif, CodeSal);
End;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 30/01/2007
Modifié le ... :   /  /    
Description .. : Chargement de la fiche
Mots clefs ... : 
*****************************************************************}
procedure TOM_EXCEPTPRESENCESAL.OnLoadRecord ;
begin
  Inherited ;
  if ds.state <> dsinsert then trttypeaffect(Nil);


   // si mode = Création et double click sur le salarié : affiche le salarié
  if (mode = 'ACTION=CREATION') and (codesal <> '') then
  begin
   setfield('PYE_SALARIE', codesal);
   setfield('NOMSAL', nomsal);
   setcontrolenabled('PYE_SALARIE',false);
   setfield('PYE_DATEDEBUT', StrToDate(datedeb));
   setfield('PYE_DATEFIN', StrToDate(datefin));
  end;
  // si mode = Modification : Initialisise les zones selon les données de l'exception
  if mode = 'ACTION=MODIFICATION'  then TrtTypeaffect(nil);

end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 26/01/2007
Modifié le ... :   /  /    
Description .. : On Argument
Mots clefs ... : 
*****************************************************************}
procedure TOM_EXCEPTPRESENCESAL.OnArgument ( S: String ) ;
var
  COMBOtypeaffect: THValComboBox;
begin
  Inherited ;
  mode := readtokenst(S);
  codesal := readtokenst(S);
  libelle := readtokenst(S);
  prenom := readtokenst(S);
  datedeb := readtokenst(S);
  datefin := readtokenst(S);
  nomsal := libelle + ' ' + prenom;


  COMBOtypeaffect := THValComboBox(GetControl('PYE_TYPEAFFECT'));
  If COMBOtypeaffect <> Nil then
  begin
    COMBOtypeaffect.OnExit := TrtTypeaffect;
    COMBOtypeaffect.ONclick := TrtTypeaffect;
  end;


end ;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 26/01/2007
Modifié le ... :   /  /
Description .. : Initialisation de la tablette du cycle affecté (journée type,
Suite ........ : modèle cycle, cycle)
Mots clefs ... : 
*****************************************************************}
procedure TOM_EXCEPTPRESENCESAL.TrtTypeaffect(sender : TObject);
begin


  if THValComboBox(GetControl('PYE_TYPEAFFECT')).value = 'CYC' then
  setcontrolproperty('PYE_CYCLEAFFECT','Datatype','PGCYCLE')
  else
  if THValComboBox(GetControl('PYE_TYPEAFFECT')).value = 'JOU' then
  setcontrolproperty('PYE_CYCLEAFFECT','Datatype','PGJOURNEETYPE')
  else
  if  THValComboBox(GetControl('PYE_TYPEAFFECT')).value = 'MOD' then
  setcontrolproperty('PYE_CYCLEAFFECT','Datatype','PGMODELECYCLE');
end;

Initialization
  registerclasses ( [ TOM_EXCEPTPRESENCESAL ] ) ;
end.
