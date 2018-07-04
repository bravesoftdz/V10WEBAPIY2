{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. : TOM gestion de la saisie des conventions collectives
Mots clefs ... :  PAIE
*****************************************************************
PT1   : 30/11/2001 SB V563 Fiche de bug n°353 Acces predefini aprés création
PT2   : 18/02/2003 SB V595 FQ 10531 Anomalie sur controle existence rubrique
                           dans table commune
PT3   : 07/04/2003 SB V_42 FQ 10609 Suite Transformation de la fiche liste en
                           mul + fiche
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par
        PgRendNoDossier() *****
PT4   : 04/10/2004 PH V_50 Suppression rechargement tablette
PT5   : 07/10/2005 VG V_60 Adaptation cahier des charges DADS-U V8R02
PT6   : 24/10/2005 VG V_60 Si on force le nouveau champ IDCC (car non renseigné)
                           on passe la fiche en modification
PT7   : 08/03/2007 PH V_70 FQ 13886 Accès CCN en STD
PT8   : 01/10/2007 FC V_80 Concepts
}
unit UTomConventionColl;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, FichList,
  {$ELSE}
  UTOB, eFichList,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOM,
  PgOutils,PgOutils2;

type
  TOM_ConventionColl = class(TOM)
  private
    LectureSeule, CEG, STD, DOS, Error, ExisteCod: Boolean;
    NomChamp: array[1..1] of Hstring;
    ValChamp: array[1..1] of variant;
    DerniereCreate: string; //PT3
    OnFerme: Boolean; //PT3
  public
    procedure OnChangeField(F: TField); override;
    procedure OnLoadRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnAfterUpdateRecord; override;
  end;

implementation
uses P5Def;
{ TOM_ConventionColl }

procedure TOM_ConventionColl.OnAfterUpdateRecord;
begin
  inherited;
  if OnFerme then Ecran.Close; //PT3
end;

procedure TOM_ConventionColl.OnChangeField(F: TField);
var
Pred, vide, Convention, mes, TempConv: string;
Ok: Boolean;
icode: integer;
begin
inherited;
vide := '';
if (F.FieldName = 'PCV_CONVENTION') then
   begin
   Convention := Getfield('PCV_CONVENTION');
   if (Convention = '') then
      exit;
   if ((isnumeric(Convention)) and (Convention <> '    ')) then
      begin
      iCode := strtoint(trim(Convention));
      TempConv := ColleZeroDevant(iCode, 3);
      if (DS.State = dsinsert) and (TempConv <> '') and
         (GetField('PCV_PREDEFINI') <> '') then
         begin
         OK := TestRubriqueCegStd(GetField('PCV_PREDEFINI'), TempConv, 0);
         if (Ok = False) then
            begin
            Mes := MesTestRubrique('CONV', GetField('PCV_PREDEFINI'), 0);
            HShowMessage('2;Code Erroné: ' + TempConv + ' ;' + mes + ';W;O;O;;;', '', '');
            TempConv := '';
            end;
         end;
      if TempConv <> Convention then
         begin
         SetField('PCV_CONVENTION', TempConv);
         SetFocusControl('PCV_CONVENTION');
         end;
      end;
   end;

if (F.FieldName = 'PCV_PREDEFINI') and (DS.State = dsinsert) then
   begin
   Pred := GetField('PCV_PREDEFINI');
   Convention := (GetField('PCV_CONVENTION'));
   if Pred = '' then
      exit;
   AccesPredefini('TOUS', CEG, STD, DOS);
{$IFDEF CPS1} // PT7 
   STD := TRUE;
{$ENDIF}
   if (Pred = 'CEG') and (CEG = FALSE) then
      begin
      Error := True;
      PGIBox('Vous ne pouvez créer de convention prédéfini CEGID', 'Accès refusé');
      Pred := '';
      SetControlProperty('PCV_PREDEFINI', 'Value', Pred);
      end;
   if (Pred = 'STD') and (STD = FALSE) then
      begin
      Error := True;
      PGIBox('Vous ne pouvez créer de convention prédéfini Standard', 'Accès refusé');
      Pred := '';
      SetControlProperty('PCV_PREDEFINI', 'Value', Pred);
      end;
   if (Convention <> '') and (Pred <> '') then
      begin
      OK := TestRubriqueCegStd(Pred, Convention, 0);
      if (Ok = False) then
         begin
         Mes := MesTestRubrique('CONV', Pred, 0);
         HShowMessage('2;Code Erroné: ' + Convention + ' ;' + mes + ';W;O;O;;;', '', '');
         SetField('PCV_CONVENTION', vide);
         if Pred <> GetField('PCV_PREDEFINI') then
            SetField('PCV_PREDEFINI', pred);
         SetFocusControl('PCV_PREDEFINI');
         exit;
         end;
      end;
   if Pred <> GetField('PCV_PREDEFINI') then
      SetField('PCV_PREDEFINI', pred);
   end;

//PT5
if (F.FieldName=('PCV_IDCC')) then
   begin
   if (GetControlText('PCV_IDCC')='') then
      begin
//PT6
      if not (ds.state in [dsinsert,dsedit]) then
         ds.edit;
//FIN PT6         
      SetControlText ('PCV_IDCC', '9999');
      end;
   SetControlText ('LIDCC',
                  RechDom ('PGIDCC', GetControlText('PCV_IDCC'), FALSE, ''));
   end;
//FIN PT5

if (ds.state in [dsBrowse]) then
   begin
   if LectureSeule then
      begin
      PaieLectureSeule(TFFicheListe(Ecran), True);
      SetControlEnabled('BDelete', False);
      end;
   SetControlEnabled('BInsert', True);
   if (CEG = FALSE) and (STD = FALSE) then
      SetControlEnabled('BInsert', False);
   SetControlEnabled('PCV_PREDEFINI', False);
   SetControlEnabled('PCV_CONVENTION', False);
   end;
end;

procedure TOM_ConventionColl.OnDeleteRecord;
begin
  inherited;
  { PT2 Anomalie Test existence rubrique : utilisation du ##PREDEFINI## }
  NomChamp[1] := '##PMI_PREDEFINI## PMI_CONVENTION';
  ValChamp[1] := GetField('PCV_CONVENTION');
  ExisteCod := RechEnrAssocier('MINIMUMCONVENT', NomChamp, ValChamp);
  if ExisteCod = TRUE then
  begin
    LastError := 1;
    LastErrorMsg := 'Attention! Certaines tables dossier utilise cette convention,' +
      '#13#10 Vous ne pouvez la supprimer!';
  end;
  if ExisteCod = False then
  begin
    NomChamp[1] := 'PSA_CONVENTION';
    ExisteCod := RechEnrAssocier('SALARIES', NomChamp, Valchamp);
    if ExisteCod = TRUE then
    begin
      LastError := 1;
      LastErrorMsg := 'Attention! Cette convention est affectée à certains salariés,' +
        '#13#10 Vous ne pouvez la supprimer!';
    end;
  end;
// PT4  ChargementTablette(TFFicheListe(Ecran).TableName, ''); //recharge les tablettes
end;

procedure TOM_ConventionColl.OnLoadRecord;
begin
  inherited;
  if not (DS.State in [dsInsert]) then DerniereCreate := ''; //PT3
  AccesPredefini('TOUS', CEG, STD, DOS);
{$IFDEF CPS1} //PT7
   STD := TRUE;
{$ENDIF}
  
  LectureSeule := FALSE;
  if (Getfield('PCV_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFicheListe(Ecran), (CEG = False));
    SetControlEnabled('BDelete', CEG);
  end;
  if (Getfield('PCV_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFicheListe(Ecran), (STD = False));
    SetControlEnabled('BDelete', STD);
  end;

  //DEB PT8
  if (CEG = False) and (STD = False) then
    SetControlEnabled('BInsert', False)
  else
    SetControlEnabled('BInsert', True);
  //FIN PT8
  SetControlEnabled('BDUPLIQUER', True);
  SetControlEnabled('PCV_PREDEFINI', False);
  SetControlEnabled('PCV_CONVENTION', False);

  if (DS.State in [dsInsert]) then
  begin
    LectureSeule := FALSE;
    PaieLectureSeule(TFFicheListe(Ecran), False);
    SetControlEnabled('PCV_PREDEFINI', True);
    SetControlEnabled('PCV_CONVENTION', True);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDelete', False);
  end;
  if (CEG = FALSE) and (STD = FALSE) then SetControlEnabled('BInsert', False);
end;

procedure TOM_ConventionColl.OnNewRecord;
begin
  inherited;

  //SetField ('PCV_CONVENTION', V_PGI_env.nodossier);

end;

procedure TOM_ConventionColl.OnUpdateRecord;
var
  Predef: string;
begin
  inherited;
  { PT3 Contrôle en création pour sortir fiche : contournement bug }
  OnFerme := False;
  if (DS.State in [dsInsert]) then
    DerniereCreate := GetField('PCV_CONVENTION')
  else
    if (DerniereCreate = GetField('PCV_CONVENTION')) then OnFerme := True; // le bug arrive on se casse !!!
  if (DS.State = dsinsert) then
  begin
    if (GetField('PCV_PREDEFINI') <> 'DOS') then
      SetField('PCV_NODOSSIER', '000000')
    else
      // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
      SetField('PCV_NODOSSIER', PgRendNoDossier());
  end;

  {If (GetField('PCV_CONVENTION')='') and (Error=False) then
    Begin
    LastError:=1;
    LastErrorMsg:='Vous devez renseigner le code de la convention';
    SetFocusControl ('PCV_CONVENTION');
    End; }
  Predef := GetField('PCV_PREDEFINI');
  if (Predef <> 'CEG') and (Predef <> 'DOS') and (Predef <> 'STD') and (Error = False) then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez renseigner le champ prédéfini';
    SetFocusControl('PCV_PREDEFINI');
  end;
  if (LastError = 0) and (Error = True) then
  begin
    LastError := 1;
    LastErrorMsg := '';
  end;
  Error := False;
  if (LastError = 0) and (GetField('PCV_CONVENTION') <> '') and (GetField('PCV_LIBELLE') <> '') then
  begin
// PT4    ChargementTablette(TFFicheListe(Ecran).TableName, ''); //recharge les tablettes
    SetControlEnabled('PCV_PREDEFINI', False); //PT1
    SetControlEnabled('PCV_CONVENTION', False); //PT1
  end;
end;

initialization
  registerclasses([TOM_ConventionColl]);
end.

