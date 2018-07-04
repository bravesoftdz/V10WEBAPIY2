{***********UNITE*************************************************
Auteur  ...... : YMO
Créé le ...... : 20/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : FPECRCREPARAM ()
Description .. : Paramètres de la fiche de choix de critères
Description .. : pour la génération des frais financiers ou de régularisation
Mots clefs ... : TOF;FPECRCREPARAM
*****************************************************************}
Unit FPECRCREPARAM_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     uTOB,
     HTB97,
     AGLInit,
     Ent1,
     SaisUtil,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     fe_main,
{$ENDIF}
     ParamSoc;

Type
  TOF_FPECRCREPARAM = Class (TOF)
    EMP_JOURNAL : THValComboBox ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;

  end ;

function CPLanceFiche_ParamCRE ( vStParam: string = '') : string ;

Implementation

Function CPLanceFiche_ParamCRE ( vStParam: string = '') : string ;
begin
   Result := AGLLanceFiche('FP','FPECRCREPARAM','','', vstParam);
end;

procedure TOF_FPECRCREPARAM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_FPECRCREPARAM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_FPECRCREPARAM.OnUpdate ;
var
  T : TOB;
  DateDeb, DateFin : String;
begin

    DateDeb:=GetControlText('EMP_DATE');
    DateFin:=GetControlText('EMP_DATE_');

    If GetControlVisible('EMP_DATE') AND (StrToDate(DateDeb)>=StrToDate(DateFin)) Then
    begin
      PgiInfo('Veuillez choisir une date de fin supérieure à la date de début de période');
      THEdit(GetControl('EMP_DATE_')).SetFocus;
      exit;
    end;

    {21/03/2007 Vérification exercice}
    If ControleDate(DateFin) =2 then
    begin
      PgiInfo('La date que vous avez renseignée n''est pas dans un exercice ouvert');
      THEdit(GetControl('EMP_DATE_')).SetFocus;
      exit;
    end;

    {FQ20542  YMO 15.06.07 Vérification journal}
    If GetControlText('EMP_JOURNAL') ='' then
    begin
      PgiInfo('Veuillez renseigner un journal');
      THEdit(GetControl('EMP_JOURNAL')).SetFocus;
      exit;
    end;

    T := TOB.Create ('', nil, -1);
    T.AddChampSupValeur('EMP_DATEDEB', DateDeb);
    T.AddChampSupValeur('EMP_DATEFIN', DateFin);
    T.AddChampSupValeur('EMP_INTER', GetControlText('EMP_INTER'));
    T.AddChampSupValeur('EMP_ASSUR', GetControlText('EMP_ASSUR'));
    T.AddChampSupValeur('EMP_JOURNAL', GetControlText('EMP_JOURNAL'));

  TheTOB := T;

  Ecran.Close ;
end ;

procedure TOF_FPECRCREPARAM.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_FPECRCREPARAM.OnArgument (S : String ) ;
var
  TypeGene, DateDebG, DateFinG : String;
  lInAn, lInMois, lInJour : Word;
begin

  TypeGene := ReadTokenSt(S);
  DateDebG := ReadTokenSt(S);

  If DateDebG <>'' then
  begin
    DecodeDate(StrToDate(DateDebG),lInAn,lInMois,lInJour);
    DateFinG := DateToStr(EncodeDate(lInAn,12,31));
  end
  else
    DateFinG:=DateToStr(VH^.EnCours.Fin);

  If TypeGene='F' then {Génération des frais financiers}
  begin
    THLabel(GetControl('TCODEEMPRUNT')).Caption:='Génération des frais financiers' ;
    THLabel(GetControl('TCODEEMPRUNT')).Visible:=True ;
  end
  else                 {Régularisation de fin de période}
  begin
    THLabel(GetControl('TCONTRE_DATE1')).Caption := 'Date de référence';

    THLabel(GetControl('TCODEEMPRUNT')).Caption:='Régularisation de fin de période' ;

    TGroupBox(GetControl('FE_FRAISFINA')).Visible:=False ;
    THLabel(GetControl('TCONTRE_DATE')).Visible:=False ;
    THEdit (GetControl('EMP_DATE')).Visible:=False ;
    THLabel(GetControl('TCODEEMPRUNT')).Visible:=True ;
  end;

  EMP_JOURNAL := THValComboBox(GetControl('EMP_JOURNAL', True)) ;

  {FQ20542 FQ20543  YMO 07.06.07  Journaux de nature OD, REG, EXT}
  EMP_JOURNAL.Plus:='J_FERME="-" AND (J_NATUREJAL="OD" OR J_NATUREJAL="EXT" OR J_NATUREJAL="REG")';

  If (EMP_JOURNAL.Values.Count > 1) then
      EMP_JOURNAL.Value:=EMP_JOURNAL.Values[0];
  {16/03/07 Reprise journal paramsoc révision}
  If GetParamSocSecur('SO_CPCUTJAL','')<>'' then  {26.06.07 FQ20543 YMO Paramètre non présent en version PCL}
    SetControlText('EMP_JOURNAL', GetParamSocSecur('SO_CPCUTJAL',''));

  SetControlText('EMP_DATE', DateDebG);
  SetControlText('EMP_DATE_', DateFinG);
  SetControlText('EMP_INTER', 'X');
  SetControlText('EMP_ASSUR', 'X');

  TheTOB := nil;

  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_FPECRCREPARAM ] ) ;
end.
