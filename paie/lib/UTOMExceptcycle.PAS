{***********UNITE*************************************************
Auteur  ...... : NA
Créé le ...... : 25/06/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : EXCEPTCYCLE
                 Saisie des exceptions d'un cycle ou modèle de cycle
Mots clefs ... : TOM;EXCEPTCYCLE
*****************************************************************
PT1  09/08/2007  FLO  Recalcul automatique des compteurs lors d'une modification/suppression
}
Unit UTOMExceptcycle;

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
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM, 
     UTob ;

Type
  TOM_Exceptcycle = Class (TOM)
    procedure OnUpdateRecord             ; override ;
    procedure OnDeleteRecord             ; override ; //PT1
    procedure OnLoadRecord               ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnAfterUpdateRecord        ; override ; //PT1
    procedure OnAfterDeleteRecord        ; override ; //PT1
  private
    mode, stArgument   : String;
    TypeCycleA,CycleA  : String; //PT1
    DateModif          : TDateTime; //PT1
    procedure TrtTypecycle(sender: TObject);
    end ;

Implementation

Uses PGPresence;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 26/06/2007
Modifié le ... :   /  /    
Description .. : Enregistrement des donnnées
Mots clefs ... : 
*****************************************************************}
procedure TOM_Exceptcycle.OnUpdateRecord ;
var
cycle, lib, typecycle, journeetype : string;
Q : TQUERY;
begin
  Inherited ;
  lib := '';
  cycle := THVALCOMBOBOX(Getcontrol('PYA_CYCLE')).value;
  typecycle := THValCombobox(Getcontrol('PYA_TYPECYCLE')).value;
  journeetype := THValCombobox(Getcontrol('PYA_JOURNEETYPE')).value;

  if (typecycle <> 'CYC') and (typecycle <> 'MOD') then
  begin
     PGIBOX('Le type de cycle n''existe pas', Ecran.Caption);
     Setfocuscontrol('PYA_TYPECYCLE');
     lasterror := 1;
  end;

  if cycle = '' then
  begin
   if typecycle = 'CYC' then PGIBOX('Le rythme de travail doit être renseigné.', Ecran.Caption)
                        else PGIBOX('Le modèle de cycle doit être renseigné.', Ecran.Caption);
   Setfocuscontrol('PYA_CYCLE');
   lasterror := 1;
  end;

  if journeetype = '' then
  begin
   PGIBOX('La journée type doit être renseignée.', Ecran.Caption);
   Setfocuscontrol('PYA_JOURNEETYPE');
   lasterror := 1;
  end;

  if typecycle = 'CYC' then
  begin
    Q := Opensql('SELECT PYC_LIBELLE FROM CYCLEENTETE WHERE PYC_CYCLE = "'+cycle+'"', true);
    if not Q.EOF then lib := Q.Findfield('PYC_LIBELLE').asstring;
  end
  else
  begin
    Q := Opensql('SELECT PMY_LIBELLE FROM MODELECYCLEENT WHERE PMY_MODELECYCLE = "'+cycle+'"', true);
    if not Q.EOF then lib := Q.Findfield('PMY_LIBELLE').asstring;
  end;
  setfield('PYA_LIBELLECYCLE', lib);
  ferme(Q);

  //PT1 - Début
  If (LastError = 0) Then
  Begin
     TypeCycleA := GetField('PYA_TYPECYCLE');
     CycleA     := GetField('PYA_CYCLE');
     DateModif  := GetField('PYA_DATEEXCEPTION');
  End;
  //PT1 - Fin
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/08/2007 / PT1
Modifié le ... :   /  /    
Description .. : Suite à la création ou la modification d'un enregistrement
Mots clefs ... : 
*****************************************************************}
procedure TOM_Exceptcycle.OnAfterUpdateRecord ;
begin
  Inherited ;
     If (LastError = 0) And (PresenceCycleIsUsed(TypeCycleA,CycleA)) Then
          CompteursARecalculer(DateModif);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/08/2007 / PT1
Modifié le ... :   /  /
Description .. : Avant la suppression d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_Exceptcycle.OnDeleteRecord ;
begin
  Inherited ;
     If (LastError = 0) Then
     Begin
          TypeCycleA := GetField('PYA_TYPECYCLE');
          CycleA     := GetField('PYA_CYCLE');
          DateModif  := GetField('PYA_DATEEXCEPTION');
     End;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 09/08/2007 / PT1
Modifié le ... :   /  /
Description .. : Suite à la suppression d'un enregistrement
Mots clefs ... :
*****************************************************************}
procedure TOM_Exceptcycle.OnAfterDeleteRecord ;
begin
  Inherited ;
     If (LastError = 0) And (PresenceCycleIsUsed(TypeCycleA, CycleA)) Then
          CompteursARecalculer(DateModif);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 25/06/2007
Modifié le ... :   /  /    
Description .. : Chargement de la fiche
Mots clefs ... : 
*****************************************************************}
procedure TOM_Exceptcycle.OnLoadRecord ;
var
  stArg, Champs, Value, S : String;
begin
  Inherited ;
  if ds.state <> dsinsert then trttypecycle(Nil);


  // si mode = Modification : Initialisise les zones selon les données de l'exception
  if mode = 'MODIFICATION'  then trttypecycle(nil);

  S := stArgument;
  stArg := ReadtokenSt(S);
  while stArg <> '' do
  begin
    Champs := READTOKENPipe(stArg,'=');
    Value := stArg;
    if (Champs <> 'ACTION') and (Champs <> '') then
      SetField(Champs, Value);
    stArg := ReadtokenSt(S);
  end;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 25/06/2007
Modifié le ... :   /  /    
Description .. : On Argument
Mots clefs ... : 
*****************************************************************}
procedure TOM_Exceptcycle.OnArgument ( S: String ) ;
var
  COMBOtypeaffect: THValComboBox;
  stArg, Champs, Value : String;
begin
  Inherited ;
  stArgument := S;
  stArg := ReadtokenSt(S);
  while stArg <> '' do
  begin
    Champs := READTOKENPipe(stArg,'=');
    Value := stArg;
    if Champs = 'ACTION' then
      mode := Value;
    stArg := ReadtokenSt(S);
  end;

  COMBOtypeaffect := THValComboBox(GetControl('PYA_TYPECYCLE'));
  If COMBOtypeaffect <> Nil then
  begin
    COMBOtypeaffect.OnExit := trttypecycle;
    COMBOtypeaffect.ONclick := trttypecycle;
  end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 25/06/2007
Modifié le ... :   /  /
Description .. : Initialisation de la tablette du cycle affecté
Suite ........ : (modèle cycle, cycle)
Mots clefs ... : 
*****************************************************************}
procedure TOM_Exceptcycle.trttypecycle(sender : TObject);
var
Libchamp: THLabel;
begin

  Libchamp := THLabel(GetControl('TPYA_CYCLE'));

  if THValComboBox(GetControl('PYA_TYPECYCLE')).value = 'CYC' then
  begin
  if Libchamp <> nil then LibChamp.caption := 'Rythme de travail';
  setcontrolproperty('PYA_CYCLE','Datatype','PGCYCLE');
  end
  else

  if  THValComboBox(GetControl('PYA_TYPECYCLE')).value = 'MOD' then
  begin
  if Libchamp <> nil then Libchamp.caption := 'Modèle de cycle';
  setcontrolproperty('PYA_CYCLE','Datatype','PGMODELECYCLE');
  end;

end;

Initialization
  registerclasses ( [ TOM_Exceptcycle ] ) ;
end.
