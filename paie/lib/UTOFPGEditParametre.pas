{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 13/09/2001
Modifié le ... :   /  /
Description .. : Editon du parmétrage de la paie
Mots clefs ... : PAIE;PARAMETRE
*****************************************************************
PT1   : 18/04/2002 SB V571 Fiche de bug n°10087 : V_PGI_env.LibDossier non
                           renseigné en Mono
PT2   : 24/09/2002 SB V585 Nouvelle Tof : TOF_PGRUBRIQUE_ETAT edition des
                           profils associés aux rubriques
PT3   : 18/12/2002 SB V591 FQ 10399 Message d'erreur à appeller sur le Onexit,
                           modification du texte
PT4   : 03/02/2004 SB V50  FQ 10686 Contrôle vraisemblance
PT5   : 11/07/2005 PH V_60 FQ 11337 Controle des bornes de debut et de fin que
                           si les cumuls sont numériques
PT6   : 18/04/2007 VG V_70 Passage par un LanceEtatTob pour amélioration des
                           performances édition des "Rubriques associées aux
                           cumuls" (18s au lieu de 2mn32 pour 53 cumuls) et
                           intégration des données issues de CUMULRUBDOSSIER
PT7   : 20/04/2007 VG V_70 Passage par un LanceEtatTob pour amélioration des
                           performances édition des "Rubriques alimentant des
                           cumuls" (1mn50 au lieu de 2mn55 pour 32 pages) et
                           intégration des données issues de CUMULRUBDOSSIER
PT8   : 28/08/2007 VG V_72 Affichage des cumuls pour chaque nouvelle page
                           FQ N°14693
PT9   : 31/08/2007 VG V_72 Erreur "Invalid column name 'PCR_NATURERUB'"
                           FQ N°14692
}
unit UTOFPGEditParametre;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
  eQRS1,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} QRS1,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, UTOB,
  ParamSoc,
  HQry,
  Ed_Tools,
  HStatus;

type
  TOF_PGCUMUL_ETAT = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure Change(Sender: TObject);
    procedure OnUpdate; override;
    procedure OnClose; override ;
  private
    TobEtat : TOB;
  end;
  TOF_PGGESTASSOC = class(TOF)
  public
    procedure OnArgument(stArgument: string); override;
    procedure OnUpdate; override;
    procedure OnClose; override ;
  private
    TobEtat : TOB;
    procedure Change(Sender: TObject);
    procedure OnChangeNature(Sender: TObject);
    procedure RenommerCritere(Reinit: boolean);
    procedure CreateCumulRub;
    end;
  TOF_PGRUBRIQUE_ETAT = class(TOF) //PT2
    procedure OnArgument(Arguments: string); override;
    procedure Change(Sender: TObject);
  end;


implementation
uses PgEditOutils, PGEditOutils2,PgOutils2;
{-------------------------------------------------------------------------------
                            CUMUL ETAT
-------------------------------------------------------------------------------- }
procedure TOF_PGCUMUL_ETAT.OnArgument(Arguments: string);
var
  Cumul, Cumul_: THEdit;
  Defaut: THEdit;
  Combo: THValComboBox;
  Min, Max: string;
begin
  inherited;
  Defaut := ThEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    //  Defaut.text:=V_PGI_env.LibDossier;  //PT1 Mise en commentaire
    Defaut.text := GetParamSoc('SO_LIBELLE');

  RecupMinMaxTablette('PG', 'CUMULPAIE', 'PCL_CUMULPAIE', Min, Max);
  Cumul := THEdit(GetControl('PCL_CUMULPAIE'));
  Cumul_ := THEdit(GetControl('PCL_CUMULPAIE_'));
  if (Cumul_ <> nil) and (Cumul <> nil) then
  begin
    Cumul.text := Min;
    Cumul_.text := Max;
    Cumul.MaxLength := 2;
    Cumul_.MaxLength := 2;
    Cumul_.OnExit := Change;
    Cumul.OnExit := Change;
  end; //PT3

  Combo := ThValComboBox(getcontrol('PCL_THEMECUM'));
  if Combo <> nil then Combo.ItemIndex := 0;
  Combo := ThValComboBox(getcontrol('PCL_ALIMCUMUL'));
  if Combo <> nil then Combo.ItemIndex := 0;
  Combo := ThValComboBox(getcontrol('PCR_NATURERUB'));
  if Combo <> nil then Combo.ItemIndex := 0;
end;

procedure TOF_PGCUMUL_ETAT.Change(Sender: TObject);
var
  temp, code, CodeCum, CodeCum_: integer;
  Cumul, Cumul_: THEdit;
begin

  Cumul := THEdit(GetControl('PCL_CUMULPAIE'));
  Cumul_ := THEdit(GetControl('PCL_CUMULPAIE_'));
  if (Cumul_ <> nil) and (Cumul <> nil) then
    if ((Cumul.text <> '') and (Cumul_.Text <> '')) And (IsNumeric (Cumul.text) and IsNumeric (Cumul_.text)) then // PT5
    begin
      val(Cumul.text, temp, code);
      CodeCum := temp;
      val(Cumul_.text, temp, code);
      CodeCum_ := temp;
      if (CodeCum) > (CodeCum_) then //PT3 modification du message
        PgiBox('La borne de fin doit être supérieure à la borne de début.', TFQRS1(Ecran).caption);
      //HShowMessage('5;Incohérence :;Vous devez saisir une borne superieur!;W;O;O;O;;;','','');
    end;
end;

//PT6
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PARAMETRE
*****************************************************************}
procedure TOF_PGCUMUL_ETAT.OnUpdate;
var
Q : TQuery;
Pages : TPageControl;
StCriteres, StPages, StPlus : string;
TCRD, TCRDD, TobEtatD : TOB;
begin
inherited;
Pages:= TPageControl(GetControl('PAGES'));
StCriteres:= RecupWhereCritere(Pages);
if (StCriteres<>'') then
   StCriteres:= StCriteres+' AND ##PCL_PREDEFINI## PCL_CUMULPAIE<>"ZZZ"'
else
   StCriteres:= 'WHERE ##PCL_PREDEFINI## PCL_CUMULPAIE<>"ZZZ"';

FreeAndNil (TobEtat);
Q:= OpenSQL ('SELECT PCL_CUMULPAIE, PCL_LIBELLE, PCL_THEMECUM, PCL_TYPECUMUL,'+
             ' PCL_ALIMCUMUL, PCL_RAZCUMUL, PCL_PREDEFINI, PCL_NODOSSIER,'+
             ' PCL_ABREGE, PCL_ALIMCUMULCOT, PCR_NATURERUB, PCR_RUBRIQUE,'+
             ' PCR_CUMULPAIE, PCR_LIBELLE, PCR_SENS, PCT_RUBRIQUE,'+
             ' PCT_LIBELLE, PCT_NATURERUB, PCT_PREDEFINI, PRM_RUBRIQUE,'+
             ' PRM_LIBELLE, PRM_NATURERUB, PRM_PREDEFINI, CO_LIBELLE'+
             ' FROM CUMULPAIE'+
             ' LEFT JOIN CUMULRUBRIQUE ON'+
             '  ##PCR_PREDEFINI##PCL_CUMULPAIE=PCR_CUMULPAIE'+
             ' LEFT JOIN COTISATION ON'+
             ' PCT_RUBRIQUE=PCR_RUBRIQUE AND'+
             ' ##PCT_PREDEFINI## PCT_NATURERUB=PCR_NATURERUB'+
             ' LEFT JOIN REMUNERATION ON'+
             ' PRM_RUBRIQUE=PCR_RUBRIQUE AND'+
             ' ##PRM_PREDEFINI## PRM_NATURERUB=PCR_NATURERUB'+
             ' LEFT JOIN COMMUN ON'+
             ' CO_TYPE="NTR" AND CO_CODE=PCR_NATURERUB '+StCriteres+
             ' ORDER BY PCL_CUMULPAIE,PCR_NATURERUB,PCR_RUBRIQUE', True);
TobEtat:= Tob.Create('Les Rubriques', nil, -1);
TobEtat.LoadDetailDB('Les Rubriques', '', '', Q, False);
Ferme(Q);

//PT9
if (StCriteres<>'') then
   StPlus:= StringReplace (StCriteres, 'PCR_', 'PKC_', [rfReplaceAll]);
//FIN PT9

Q:= OpenSQL ('SELECT PCL_CUMULPAIE, PCL_LIBELLE, PCL_THEMECUM, PCL_TYPECUMUL,'+
             ' PCL_ALIMCUMUL, PCL_RAZCUMUL, PCL_PREDEFINI, PCL_NODOSSIER,'+
             ' PCL_ABREGE, PCL_ALIMCUMULCOT, PKC_NATURERUB AS PCR_NATURERUB,'+
             ' PKC_RUBRIQUE AS PCR_RUBRIQUE, PKC_CUMULPAIE AS PCR_CUMULPAIE,'+
             ' PKC_LIBELLE AS PCR_LIBELLE, PKC_SENS AS PCR_SENS, PCT_RUBRIQUE,'+
             ' PCT_LIBELLE, PCT_NATURERUB, PCT_PREDEFINI, PRM_RUBRIQUE,'+
             ' PRM_LIBELLE, PRM_NATURERUB, PRM_PREDEFINI, CO_LIBELLE'+
             ' FROM CUMULRUBDOSSIER'+
             ' LEFT JOIN CUMULPAIE ON'+
             ' PCL_CUMULPAIE=PKC_CUMULPAIE'+
             ' LEFT JOIN COTISATION ON'+
             ' PCT_RUBRIQUE=PKC_RUBRIQUE AND'+
             ' ##PCT_PREDEFINI## PCT_NATURERUB=PKC_NATURERUB'+
             ' LEFT JOIN REMUNERATION ON'+
             ' PRM_RUBRIQUE=PKC_RUBRIQUE AND'+
             ' ##PRM_PREDEFINI## PRM_NATURERUB=PKC_NATURERUB'+
             ' LEFT JOIN COMMUN ON'+
             ' CO_TYPE="NTR" AND CO_CODE=PKC_NATURERUB '+StPlus+
             ' ORDER BY PCL_CUMULPAIE,PCR_NATURERUB,PCR_RUBRIQUE', True);
TCRD:= Tob.Create('Les Rubriques', nil, -1);
TCRD.LoadDetailDB('Les Rubriques', '', '', Q, False);
Ferme(Q);

TCRDD:= TCRD.FindFirst ([''], [''], False);
While (TCRDD<>nil) do
      begin
      TobEtatD:= TobEtat.FindFirst (['PCL_CUMULPAIE', 'PCR_NATURERUB',
                                     'PCR_RUBRIQUE'],
                                    [TCRDD.GetValue ('PCL_CUMULPAIE'),
                                     TCRDD.GetValue ('PCR_NATURERUB'),
                                     TCRDD.GetValue ('PCR_RUBRIQUE')], False);
      if (Assigned (TobEtatD)) then
         begin
         if (TobEtatD.GetValue ('PCR_SENS')<>TCRDD.GetValue('PCR_SENS')) then
            begin
            FreeAndNil(TobEtatD);
            TCRDD.ChangeParent (TobEtat, -1);
            end
         else
            begin
            FreeAndNil(TobEtatD);
            FreeAndNil(TCRDD);
            end;
         end
      else
         TCRDD.ChangeParent (TobEtat, -1);

      TCRDD:= TCRD.FindFirst ([''], [''], False);
      end;
FreeAndNil (TCRD);

TobEtat.Detail.Sort('PCL_CUMULPAIE;PCR_NATURERUB;PCR_RUBRIQUE');

StPages:= AglGetCriteres (Pages, FALSE);
TFQRS1(Ecran).TypeEtat:= 'E';
TFQRS1(Ecran).NatureEtat:= 'PGA';
TFQRS1(Ecran).CodeEtat:= 'PCU';
TFQRS1(Ecran).LaTob:= TobEtat;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PARAMETRE
*****************************************************************}
procedure TOF_PGCUMUL_ETAT.OnClose ;
begin
Inherited ;
FreeAndNil (TobEtat);
end ;
//FIN PT6

{-------------------------------------------------------------------------------
                           RUBRIQUES
-------------------------------------------------------------------------------- }
{ TOF_PGGESTASSOC }

procedure TOF_PGGESTASSOC.OnArgument(stArgument: string);
var
  GA: TCheckBox;
  Defaut: THEdit;
  Combo: ThValComboBox;
begin
  inherited;
{PT7
  Grille := THGrid(GetControl('GDCUMUL'));
}
  Defaut := ThEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    //  Defaut.text:=V_PGI_env.LibDossier; PT1 Mise en commentaire
    Defaut.text := GetParamSoc('SO_LIBELLE');

  GA := TCheckBox(GetControl('CBGESASS'));
  if GA <> nil then GA.OnClick := Change;
  GA := TCheckBox(GetControl('CBCOT'));
  if GA <> nil then GA.OnClick := Change;
  GA := TCheckBox(GetControl('CBREM'));
  if GA <> nil then GA.OnClick := Change;

  combo := ThValComboBox(getcontrol('PCR_NATURERUB'));
  if combo <> nil then Combo.OnChange := OnChangeNature;
{PT7
  SetcontrolProperty('GDCUMUL', 'Visible', False);
}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... :
Créé le ...... :   /  /    
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PARAMETRE
*****************************************************************}
procedure TOF_PGGESTASSOC.Change(Sender: TObject);
var
GA, GCot, GRem: TCheckBox;
Zone: Tcontrol;
begin
GA:= TCheckBox (GetControl ('CBGESASS'));
GCot:= TCheckBox (GetControl ('CBCOT'));
GRem:= TCheckBox (GetControl ('CBREM'));
SetControlEnabled ('FLISTE', True);
if (GA <> nil) and (GCot <> nil) and (GRem <> nil) then
   begin
   if (Ga.Checked = TRUE) then
      begin
{PT7
      TFQRS1(Ecran).CodeEtat:= 'PGR';
}
      GCot.Enabled:= False;
      GRem.Enabled:= False;
      SetControlEnabled ('FLISTE', False);
      SetControlChecked ('FLISTE', False);
      end
   else
   if (Ga.Checked = FALSE) and (Ga.enabled = True) then
      begin
{PT7
      TFQRS1(Ecran).CodeEtat:= 'PGR';
}
      GCot.Enabled:= TRUE;
      GRem.Enabled:= TRUE;
      end;

   if (GCot.Checked = TRUE) then
      begin
{PT7
      TFQRS1(Ecran).CodeEtat:= 'PCT';
}
      Ga.Enabled:= False;
      GRem.Enabled:= False;
      end
   else
   if (GCot.Checked = FALSE) and (Gcot.enabled = True) then
      begin
{PT7
      TFQRS1(Ecran).CodeEtat:= 'PGR';
}
      Ga.Enabled:= TRUE;
      GRem.Enabled:= True;
      end;

   if (GRem.Checked = TRUE) then
      begin
{PT7
      TFQRS1(Ecran).CodeEtat:= 'PRM';
}
      GCot.Enabled:= False;
      Ga.Enabled:= False;
      end
   else
   if (GRem.Checked = FALSE) and (GRem.enabled = True) then
      begin
{PT7
      TFQRS1(Ecran).CodeEtat:= 'PGR';
}
      GCot.Enabled:= TRUE;
      Ga.Enabled:= True;
      end;
   end;

if (GA <> nil) then
   if Ga.Checked = TRUE then
      begin
      SetcontrolProperty ('TBRUBRIQUE', 'TabVisible', True);
{PT7
      Tob_Cumul := Tob.Create('Les cumuls', nil, -1);
      Tob_cumul.LoadDetailDB('CUMULPAIE', '', '', nil, False);
      Tob_cumul.PutGridDetail(Grille, FALSE, TRUE, 'PCL_CUMULPAIE;PCL_LIBELLE', TRUE);
      Tob_Cumul.free;
}
      Zone:= ThValComboBox (getcontrol ('PCR_NATURERUB'));
      InitialiseCombo (Zone);
      end
   else
      begin
      SetControlProperty ('PCR_NATURERUB', 'Value', '');
      SetControlProperty ('PCR_RUBRIQUE', 'Text', '');
      SetControlProperty ('PCR_RUBRIQUE_', 'Text', '');
      SetcontrolProperty ('TBRUBRIQUE', 'TabVisible', False);
      end;
end;

procedure TOF_PGGESTASSOC.OnChangeNature(Sender: TObject);
var
  GA: TCheckBox;
  Min, Max: string;
begin
  GA := TCheckBox(GetControl('CBGESASS'));
  if (GA <> nil) then
    if Ga.Checked = False then
    begin
      if GetControltext('PCR_NATURERUB') <> '' then SetControlText('PCR_NATURERUB', '');
      if GetControltext('PCR_RUBRIQUE') <> '' then SetControlText('PCR_RUBRIQUE', '');
      if GetControltext('PCR_RUBRIQUE_') <> '' then SetControlText('PCR_RUBRIQUE_', '');
      RenommerCritere(False);
      Exit;
    end;
  RenommerCritere(True);
  if GetControlText('PCR_NATURERUB') = '' then
  begin
    SetControlProperty('PCR_RUBRIQUE', 'DataType', '');
    SetControlProperty('PCR_RUBRIQUE_', 'DataType', '');
    SetControlProperty('PCR_RUBRIQUE', 'Text', '');
    SetControlProperty('PCR_RUBRIQUE_', 'Text', '');
    SetControlEnabled('PCR_RUBRIQUE', False);
    SetControlEnabled('PCR_RUBRIQUE_', False);
  end
  else
  begin
    SetControlEnabled('PCR_RUBRIQUE', True);
    SetControlEnabled('PCR_RUBRIQUE_', True);
  end;

  if GetControlText('PCR_NATURERUB') = 'AAA' then
  begin
    SetControlProperty('PCR_RUBRIQUE', 'DataType', 'PGREMUNERATION');
    SetControlProperty('PCR_RUBRIQUE_', 'DataType', 'PGREMUNERATION');
    RecupMinMaxTablette('PG', 'REMUNERATION', 'PRM_RUBRIQUE', Min, Max);
    SetControlProperty('PCR_RUBRIQUE', 'Text', Min);
    SetControlProperty('PCR_RUBRIQUE_', 'Text', Max);
  end;
  if GetControlText('PCR_NATURERUB') = 'BAS' then
  begin
    SetControlProperty('PCR_RUBRIQUE', 'DataType', 'PGBASECOTISATION');
    SetControlProperty('PCR_RUBRIQUE_', 'DataType', 'PGBASECOTISATION');
    RecupMinMaxTablette('PG', 'COTISATION', 'PCT_RUBRIQUE', Min, Max);
    SetControlProperty('PCR_RUBRIQUE', 'Text', Min);
    SetControlProperty('PCR_RUBRIQUE_', 'Text', Max);
  end;
  if GetControlText('PCR_NATURERUB') = 'COT' then
  begin
    SetControlProperty('PCR_RUBRIQUE', 'DataType', 'PGCOTIS'); //PT4
    SetControlProperty('PCR_RUBRIQUE_', 'DataType', 'PGCOTIS');
    RecupMinMaxTablette('PG', 'COTISATION', 'PCT_RUBRIQUE', Min, Max);
    SetControlProperty('PCR_RUBRIQUE', 'Text', Min);
    SetControlProperty('PCR_RUBRIQUE_', 'Text', Max);
  end;
end;

procedure TOF_PGGESTASSOC.RenommerCritere(Reinit: boolean);
begin
  if Reinit = False then
  begin
    SetControlProperty('PCR_NATURERUB', 'Name', 'NATURERUB');
    SetControlProperty('PCR_RUBRIQUE', 'Name', 'RUBRIQUE');
    SetControlProperty('PCR_RUBRIQUE_', 'Name', 'RUBRIQUE_');
  end
  else
  begin
    SetControlProperty('NATURERUB', 'Name', 'PCR_NATURERUB');
    SetControlProperty('RUBRIQUE', 'Name', 'PCR_RUBRIQUE');
    SetControlProperty('RUBRIQUE_', 'Name', 'PCR_RUBRIQUE_');
  end;
end;

//PT7
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 18/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PARAMETRE
*****************************************************************}
procedure TOF_PGGESTASSOC.OnUpdate;
var
Q : TQuery;
Pages : TPageControl;
StCriteres, StPages, StPlus : string;
TCRD, TCRDD, TobEtatD : TOB;
i, MaxCum : integer;
begin
inherited;
Pages:= TPageControl(GetControl('PAGES'));
if ((GetCheckBoxState ('CBGESASS')=CbChecked) or
   (GetControlEnabled ('CBGESASS')=True)) then
   begin
   StCriteres:= RecupWhereCritere(Pages);
   if (StCriteres<>'') then
      StPlus:= StCriteres+' AND ##PCR_PREDEFINI## PCR_NATURERUB<>"ZZZ"'
   else
      StPlus:= 'WHERE ##PCR_PREDEFINI## PCR_NATURERUB<>"ZZZ"';

   FreeAndNil (TobEtat);
   Q:= OpenSQL ('SELECT DISTINCT PCR_NATURERUB, PCR_RUBRIQUE, PRM_ABREGE,'+
                ' PCT_ABREGE, PCT_PREDEFINI, PRM_PREDEFINI, CO_LIBELLE'+
                ' FROM CUMULRUBRIQUE'+
                ' LEFT JOIN COTISATION ON'+
                ' PCR_NATURERUB=PCT_NATURERUB AND'+
                ' ##PCT_PREDEFINI## PCR_RUBRIQUE=PCT_RUBRIQUE'+
                ' LEFT JOIN REMUNERATION ON'+
                ' PCR_NATURERUB=PRM_NATURERUB AND'+
                ' ##PRM_PREDEFINI## PCR_RUBRIQUE=PRM_RUBRIQUE'+
                ' LEFT JOIN COMMUN ON'+
                ' CO_TYPE="NTR" AND CO_CODE=PCR_NATURERUB '+StPlus+
                ' ORDER BY PCR_NATURERUB, PCR_RUBRIQUE', True);
   TobEtat:= Tob.Create('Les Rubriques', nil, -1);
   TobEtat.LoadDetailDB('Les Rubriques', '', '', Q, False);
   Ferme(Q);

   if (StCriteres<>'') then
      begin
      StPlus:= StringReplace (StCriteres, 'PCR_', 'PKC_', [rfReplaceAll]);
      StPlus:= StPlus+' AND PKC_NATURERUB<>"ZZZ"';
      end
   else
      StPlus:= 'WHERE PKC_NATURERUB<>"ZZZ"';

   Q:= OpenSQL ('SELECT DISTINCT PKC_NATURERUB AS PCR_NATURERUB,'+
                ' PKC_RUBRIQUE AS PCR_RUBRIQUE, PRM_ABREGE, PCT_ABREGE,'+
                ' PCT_PREDEFINI, PRM_PREDEFINI, CO_LIBELLE'+
                ' FROM CUMULRUBDOSSIER'+
                ' LEFT JOIN COTISATION ON'+
                ' PKC_NATURERUB=PCT_NATURERUB AND'+
                ' ##PCT_PREDEFINI## PKC_RUBRIQUE=PCT_RUBRIQUE'+
                ' LEFT JOIN REMUNERATION ON'+
                ' PKC_NATURERUB=PRM_NATURERUB AND'+
                ' ##PRM_PREDEFINI## PKC_RUBRIQUE=PRM_RUBRIQUE'+
                ' LEFT JOIN COMMUN ON'+
                ' CO_TYPE="NTR" AND CO_CODE=PKC_NATURERUB '+StPlus+
                ' ORDER BY PCR_NATURERUB, PCR_RUBRIQUE', True);
   TCRD:= Tob.Create('Les Rubriques', nil, -1);
   TCRD.LoadDetailDB('Les Rubriques', '', '', Q, False);
   Ferme(Q);

   TCRDD:= TCRD.FindFirst ([''], [''], False);
   While (TCRDD<>nil) do
         begin
         TobEtatD:= TobEtat.FindFirst (['PCR_NATURERUB', 'PCR_RUBRIQUE'],
                                       [TCRDD.GetValue ('PCR_NATURERUB'),
                                        TCRDD.GetValue ('PCR_RUBRIQUE')], False);
         if (Assigned (TobEtatD)) then
            begin
            FreeAndNil(TCRDD);
            end
         else
            TCRDD.ChangeParent (TobEtat, -1);

         TCRDD:= TCRD.FindFirst ([''], [''], False);
         end;
   FreeAndNil (TCRD);

   if (TobEtat.Detail.Count=0) then
      FreeAndNil (TobEtat);

   Q:= OpenSql ('SELECT COUNT (*) NBRE'+
                ' FROM CUMULPAIE WHERE'+
                ' PCL_CUMULPAIE <> "" AND ##PCL_PREDEFINI##', True);
   if (not Q.EOF) then
      MaxCum:= Q.FindField ('NBRE').asInteger
   else
      exit;
   Ferme(Q);
   TobEtatD:= TobEtat.FindFirst ([''], [''], False);
   TobEtatD.AddChampSupValeur ('TYPEINFO', 'LIBELLE', True);
   for i:= 1 to MaxCum do
       TobEtatD.AddChampSup ('C'+IntToStr (i), True);

{PT8
   TobEtatD:= TobEtat.FindFirst (['PCR_NATURERUB'], ['AAA'], False);
   if (TobEtatD<>nil) then
      begin
      TobEtatD:= Tob.Create ('Les Rubriques', TobEtat, -1);
      TobEtatD.AddChampSupValeur ('PCR_NATURERUB', 'AAA');
      TobEtatD.AddChampSupValeur ('PCR_RUBRIQUE', '0000');
      TobEtatD.AddChampSupValeur ('PRM_ABREGE', '');
      TobEtatD.AddChampSupValeur ('PCT_ABREGE', '');
      TobEtatD.AddChampSupValeur ('PCT_PREDEFINI', '');
      TobEtatD.AddChampSupValeur ('PRM_PREDEFINI', '');
      TobEtatD.AddChampSupValeur ('CO_LIBELLE', 'Rémunération');
      TobEtatD.AddChampSupValeur ('TYPEINFO', 'Rémunération');
      for i:= 1 to MaxCum do
          TobEtatD.AddChampSup ('C'+IntToStr (i), False);
      end;

   TobEtatD:= TobEtat.FindFirst (['PCR_NATURERUB'], ['BAS'], False);
   if (TobEtatD<>nil) then
      begin
      TobEtatD:= Tob.Create ('Les Rubriques', TobEtat, -1);
      TobEtatD.AddChampSupValeur ('PCR_NATURERUB', 'BAS');
      TobEtatD.AddChampSupValeur ('PCR_RUBRIQUE', '0000');
      TobEtatD.AddChampSupValeur ('PRM_ABREGE', '');
      TobEtatD.AddChampSupValeur ('PCT_ABREGE', '');
      TobEtatD.AddChampSupValeur ('PCT_PREDEFINI', '');
      TobEtatD.AddChampSupValeur ('PRM_PREDEFINI', '');
      TobEtatD.AddChampSupValeur ('CO_LIBELLE', 'Base de cotisation');
      TobEtatD.AddChampSupValeur ('TYPEINFO', 'Base de cotisation');
      for i:= 1 to MaxCum do
          TobEtatD.AddChampSup ('C'+IntToStr (i), False);
      end;

   TobEtatD:= TobEtat.FindFirst (['PCR_NATURERUB'], ['COT'], False);
   if (TobEtatD<>nil) then
      begin
      TobEtatD:= Tob.Create ('Les Rubriques', TobEtat, -1);
      TobEtatD.AddChampSupValeur ('PCR_NATURERUB', 'COT');
      TobEtatD.AddChampSupValeur ('PCR_RUBRIQUE', '0000');
      TobEtatD.AddChampSupValeur ('PRM_ABREGE', '');
      TobEtatD.AddChampSupValeur ('PCT_ABREGE', '');
      TobEtatD.AddChampSupValeur ('PCT_PREDEFINI', '');
      TobEtatD.AddChampSupValeur ('PRM_PREDEFINI', '');
      TobEtatD.AddChampSupValeur ('CO_LIBELLE', 'Cotisation');
      TobEtatD.AddChampSupValeur ('TYPEINFO', 'Cotisation');
      for i:= 1 to MaxCum do
          TobEtatD.AddChampSup ('C'+IntToStr (i), False);
      end;
}

   TobEtat.Detail.Sort('PCR_NATURERUB;PCR_RUBRIQUE;TYPEINFO');

   CreateCumulRub;
   end;

StPages:= AglGetCriteres (Pages, FALSE);
TFQRS1(Ecran).TypeEtat:= 'E';
TFQRS1(Ecran).NatureEtat:= 'PGA';

if ((GetCheckBoxState ('CBGESASS')=CbChecked) or
   (GetControlEnabled ('CBGESASS')=True)) then
   begin
   TFQRS1(Ecran).CodeEtat:= 'PGR';
   TFQRS1(Ecran).LaTob:= TobEtat;
   end
else
if (GetCheckBoxState ('CBCOT')=CbChecked) then
   TFQRS1(Ecran).CodeEtat:= 'PCT'
else
if (GetCheckBoxState ('CBREM')=CbChecked) then
   TFQRS1(Ecran).CodeEtat:= 'PRM';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 18/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PARAMETRE
*****************************************************************}
procedure TOF_PGGESTASSOC.OnClose ;
begin
Inherited ;
FreeAndNil (TobEtat);
end ;
//FIN PT6

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 19/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PARAMETRE
*****************************************************************}
procedure TOF_PGGESTASSOC.CreateCumulRub;
var
Q : TQuery;
TCR, TCRD, TCRF, TCRDF, Tob_Cumul, Tob_CumulD, TobEtatD : TOB;
i : Integer;
begin
Tob_Cumul:= Tob.Create('Les cumuls', nil, -1);
Tob_cumul.LoadDetailDB ('CUMULPAIE', '', '', nil, False);

Q:= OpenSQL ('SELECT DISTINCT PCR_NATURERUB, PCR_RUBRIQUE, PCR_CUMULPAIE,'+
             ' PCR_LIBELLE, PCR_SENS'+
             ' FROM CUMULRUBRIQUE WHERE'+
             ' ##PCR_PREDEFINI##'+
             ' ORDER BY PCR_NATURERUB, PCR_RUBRIQUE, PCR_CUMULPAIE', True);
TCR:= Tob.Create('La Tob chargee', nil, -1);
TCR.LoadDetailDB('Mon CUMULRUBRIQUE', '', '', Q, False);
Ferme(Q);

Q:= OpenSQL ('SELECT DISTINCT PKC_NATURERUB AS PCR_NATURERUB,'+
             ' PKC_RUBRIQUE AS PCR_RUBRIQUE, PKC_CUMULPAIE AS PCR_CUMULPAIE,'+
             ' PKC_LIBELLE AS PCR_LIBELLE, PKC_SENS AS PCR_SENS'+
             ' FROM CUMULRUBDOSSIER'+
             ' ORDER BY PCR_NATURERUB, PCR_RUBRIQUE, PCR_CUMULPAIE', True);
TCRD:= Tob.Create('La Tob chargee', nil, -1);
TCRD.LoadDetailDB('Mon CUMULRUBRIQUE', '', '', Q, False);
Ferme(Q);

TCRDF:= TCRD.FindFirst ([''], [''], False);
While (TCRDF<>nil) do
      begin
      TCRF:= TCR.FindFirst (['PCR_NATURERUB', 'PCR_RUBRIQUE', 'PCR_CUMULPAIE'],
                            [TCRDF.GetValue ('PCR_NATURERUB'),
                             TCRDF.GetValue ('PCR_RUBRIQUE'),
                             TCRDF.GetValue ('PCR_CUMULPAIE')], False);
      if (Assigned (TCRF)) then
         begin
         if (TCRF.GetValue ('PCR_SENS')<>TCRDF.GetValue('PCR_SENS')) then
            begin
            FreeAndNil(TCRF);
            TCRDF.ChangeParent (TCR, -1);
            end
         else
            begin
            FreeAndNil(TCRF);
            FreeAndNil(TCRDF);
            end;
         end
      else
         TCRDF.ChangeParent (TCR, -1);

      TCRDF:= TCRD.FindFirst ([''], [''], False);
      end;
FreeAndNil (TCRD);

TCR.Detail.Sort('PCR_NATURERUB;PCR_RUBRIQUE;PCR_CUMULPAIE');

InitMoveProgressForm (NIL,'Chargement des données',
                      'Veuillez patienter SVP ...', TobEtat.Detail.Count, False,
                      True);

TobEtatD:= TobEtat.FindFirst ([''], [''], False);
While (TobEtatD<>nil) do
      begin
{PT8
      if (TobEtatD.GetValue ('PCR_RUBRIQUE')='0000') then
         begin
         for i:= 1 to Tob_Cumul.Detail.Count do
             TobEtatD.PutValue ('C'+IntToStr (i),
                                Tob_Cumul.Detail[i-1].GetValue('PCL_CUMULPAIE'));
         end;
}
      for i:= 1 to Tob_Cumul.Detail.Count do
          SetControlText ('C'+IntToStr (i),
                          Tob_Cumul.Detail[i-1].GetValue('PCL_CUMULPAIE'));
//FIN PT8

      if (TCR <> nil) then
         TCRF:= TCR.FindFirst (['PCR_NATURERUB', 'PCR_RUBRIQUE'],
                               [TobEtatD.GetValue ('PCR_NATURERUB'),
                                TobEtatD.GetValue ('PCR_RUBRIQUE')], FALSE);
      if (TCRF <> nil) then
         begin
         while TCRF <> nil do
               begin
               Tob_CumulD:= Tob_Cumul.FindFirst (['PCL_CUMULPAIE'],
                                                 [TCRF.GetValue ('PCR_CUMULPAIE')], False);
               i:= Tob_CumulD.GetIndex;
               TobEtatD.PutValue ('C'+IntToStr (i+1), TCRF.GetValue ('PCR_SENS'));
               TCRF:= TCR.FindNext (['PCR_NATURERUB', 'PCR_RUBRIQUE'],
                                    [TobEtatD.GetValue ('PCR_NATURERUB'),
                                     TobEtatD.GetValue ('PCR_RUBRIQUE')], FALSE);
               end;
         end;
      MoveCurProgressForm (TobEtatD.GetValue ('PCR_RUBRIQUE'));
      TobEtatD:= TobEtat.FindNext ([''], [''], False);
      end;
FiniMoveProgressForm;

FreeAndNil (Tob_Cumul);
FreeAndNil (TCR);
end;
//FIN PT7


{ TOF_PGRUBRIQUE_ETAT }
//DEB PT2
procedure TOF_PGRUBRIQUE_ETAT.OnArgument(Arguments: string);
var
  combo: ThValComboBox;
begin
  inherited;
  SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
  combo := ThValComboBox(getcontrol('PPM_NATURERUB'));
  if combo <> nil then Combo.OnChange := Change;

end;

procedure TOF_PGRUBRIQUE_ETAT.Change(Sender: TObject);
var
  Min, Max: string;
begin
  if GetControlText('PPM_NATURERUB') = '' then
  begin
    SetControlProperty('PPM_RUBRIQUE', 'DataType', '');
    SetControlProperty('PPM_RUBRIQUE_', 'DataType', '');
    SetControlProperty('PPM_RUBRIQUE', 'Text', '');
    SetControlProperty('PPM_RUBRIQUE_', 'Text', '');
    SetControlEnabled('PPM_RUBRIQUE', False);
    SetControlEnabled('PPM_RUBRIQUE_', False);
  end
  else
    if GetControlText('PPM_NATURERUB') = 'AAA' then
  begin
    SetControlEnabled('PPM_RUBRIQUE', True);
    SetControlEnabled('PPM_RUBRIQUE_', True);
    SetControlProperty('PPM_RUBRIQUE', 'DataType', 'PGREMUNERATION');
    SetControlProperty('PPM_RUBRIQUE_', 'DataType', 'PGREMUNERATION');
    RecupMinMaxTablette('PG', 'REMUNERATION', 'PRM_RUBRIQUE', Min, Max);
    SetControlProperty('PPM_RUBRIQUE', 'Text', Min);
    SetControlProperty('PPM_RUBRIQUE_', 'Text', Max);
  end
  else
    if GetControlText('PPM_NATURERUB') = 'BAS' then
  begin
    SetControlEnabled('PPM_RUBRIQUE', True);
    SetControlEnabled('PPM_RUBRIQUE_', True);
    SetControlProperty('PPM_RUBRIQUE', 'DataType', 'PGBASECOTISATION');
    SetControlProperty('PPM_RUBRIQUE_', 'DataType', 'PGBASECOTISATION');
    RecupMinMaxTablette('PG', 'COTISATION', 'PCT_RUBRIQUE', Min, Max);
    SetControlProperty('PPM_RUBRIQUE', 'Text', Min);
    SetControlProperty('PPM_RUBRIQUE_', 'Text', Max);
  end
  else
    if GetControlText('PPM_NATURERUB') = 'COT' then
  begin
    SetControlEnabled('PPM_RUBRIQUE', True);
    SetControlEnabled('PPM_RUBRIQUE_', True);
    SetControlProperty('PPM_RUBRIQUE', 'DataType', 'PGCOTISATION');
    SetControlProperty('PPM_RUBRIQUE_', 'DataType', 'PGCOTISATION');
    RecupMinMaxTablette('PG', 'COTISATION', 'PCT_RUBRIQUE', Min, Max);
    SetControlProperty('PPM_RUBRIQUE', 'Text', Min);
    SetControlProperty('PPM_RUBRIQUE_', 'Text', Max);
  end;
end;
//FIN PT2


initialization
  registerclasses([TOF_PGCUMUL_ETAT, TOF_PGGESTASSOC, TOF_PGRUBRIQUE_ETAT]);

end.

