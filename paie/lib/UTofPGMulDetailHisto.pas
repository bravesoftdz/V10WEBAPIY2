{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 23/05/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULDETAILHISTO ()
Mots clefs ... : TOF;PGMULDETAILHISTO
*****************************************************************
PT1   21/06/2007 FC V_72 FQ 14330 Accès à tort aux salariés confidentiels
}
Unit UTofPGMulDetailHisto ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
db,     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,Mul,HDB,
{$else}
     eMul,

     MainEAGL,
{$ENDIF}
     forms,
     uTob,
     HQry,
     sysutils,
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HTB97,
     UTOF,
     StrUtils,EntPaie, PGOutils2 ;

Type
  TOF_PGMULDETAILHISTO = Class (TOF)
      procedure OnClose                  ; override ;
      procedure OnArgument (S : String ) ; override ;
      procedure OnLoad                   ; override ;
    private
      procedure ActiveWhere(Sender: TObject);
      procedure GrilleDblClick(Sender : TObject);
      procedure CreerHisto(Sender : TObject);
      procedure OnClickSalarieSortie(Sender: TObject);
      Procedure PGRechSal (FF : TForm) ;
      procedure ExitEdit(Sender: TObject);
  end ;
var PGTobSalMul,PGTobTableDynMul : Tob;

Implementation

procedure TOF_PGMULDETAILHISTO.OnClose;
begin
  Inherited ;
  If Assigned(PGTobSalMul) then FreeANdNil(PGTobSalMul);
  If Assigned(PGTobTableDynMul) then FreeANdNil(PGTobTableDynMul);
end;

procedure TOF_PGMULDETAILHISTO.OnArgument (S : String ) ;
var    {$IFNDEF EAGLCLIENT}
        Liste : THDBGrid ;
        {$ELSE}
        Liste : THGrid ;
        {$ENDIF}
        BNew : TToolBarButton97;
        Q : TQuery;
        Check : TCheckBox;
begin
  Inherited ;
            {$IFNDEF EAGLCLIENT}
        Liste := THDBGrid(GetControl('FListe'));
        {$ELSE}
        Liste := THGrid(GetControl('FListe'));
        {$ENDIF}
        If Liste <> Nil Then Liste.OnDblClick := GrilleDblClick ;
        BNew := TToolBarButton97(Getcontrol('BInsert'));
       If BNew <> Nil then BNew.OnClick := CreerHisto;
       Q := OpenSQL('SELECT PSA_SALARIE,PSA_CONVENTION,PSA_ETABLISSEMENT FROM SALARIES',True);
        PGTobSalMul := Tob.Create('LeSalarie',Nil,-1);
        PGTobSalMul.LoadDetailDB('LeSalarie','','',Q,False);
        Ferme(Q);
        Q := OpenSQL('SELECT * FROM TABLEDIMDET ORDER BY PTD_DTVALID DESC',True);
        PGTobTableDynMul := Tob.Create('ParamTableDyn',Nil,-1);
        PGTobTableDynMul.LoadDetailDB('ParamTableDyn','','',Q,False);
        Ferme(Q);

        Check := TCheckBox(GetControl('CKSORTIE'));
        if (Check <> nil) then
          Check.OnClick:=OnClickSalarieSortie;

  PGRechSal (Ecran);
end ;

procedure TOF_PGMULDETAILHISTO.GrilleDblClick(Sender : TObject);
var Code,Retour : String;
    Q_Mul : THQuery ;
begin
     {$IFDEF EAGLCLIENT}
     TFmul(Ecran).Q.TQ.Seek(TFmul(Ecran).FListe.Row-1) ;  //PT1
     {$ENDIF}
     Q_Mul := THQuery(Ecran.FindComponent('Q')) ;
     Code := Q_Mul.FindField('PHD_GUIDHISTO').AsString;
     Retour := AGLLanceFiche('PAY','PGHISTODETAIL','',Code,'ACTION=CONSULTATION');
     If Retour <> '' then TFMul(Ecran).BChercheClick(TFMul(Ecran).BCHerche);
end;

procedure TOF_PGMULDETAILHISTO.CreerHisto(Sender : TObject);
Var Retour : String;
begin
     Retour := AGLLanceFiche('PAY','PGHISTODETAIL','','',';;ACTION=CREATION');
     If Retour <> '' then TFMul(Ecran).BChercheClick(TFMul(Ecran).BCHerche);
end;

procedure TOF_PGMULDETAILHISTO.ActiveWhere(Sender: TObject);
var
  StWhere,Info : String;
  DateArret,DateApplic : TDateTime;
  StDateArret : String;
  i : integer;
  St : String;
begin
  SetControlText('XX_WHERE','');
  
  StWhere := ' AND PHD_PGTYPEINFOLS="SAL"';

  // Critère matricule salarié
  if (GetControlText('PSA_SALARIE') <> '') then
    stWhere := stWhere + ' AND PSA_SALARIE = "' + GetControlText('PSA_SALARIE') + '"';

  // Critère nom salarié
  if (GetControlText('PSA_LIBELLE') <> '') then
    stWhere := stWhere + ' AND PSA_LIBELLE LIKE "' + GetControlText('PSA_LIBELLE') + '%"';

  // Critère établissement
  if (GetControlText('PSA_ETABLISSEMENT') <> '') and (GetControlText('PSA_ETABLISSEMENT') <> '<<Tous>>') then
    stWhere := stWhere + ' AND PSA_ETABLISSEMENT = "' + GetControlText('PSA_ETABLISSEMENT') + '"';

  // Information modifiée
  if (GetControlText('PHD_PGINFOSMODIF') <> '') and (GetControlText('PHD_PGINFOSMODIF') <> '<<Tous>>') then
  begin
    Info := GetControlText('PHD_PGINFOSMODIF');
    stWhere := stWhere + ' AND (';
    while Info <> '' do
    begin
      stWhere := stWhere + ' PHD_PGINFOSMODIF = "' + READTOKENST(Info) + '"';
      if (Info <> '') then
        stWhere := stWhere + ' OR ';
    end;
    stWhere := stWhere + ')';
  end;

  // Gérer la case à cocher exclure
  if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET'))) then
  Begin
    DateArret := StrtoDate(GetControlText('DATEARRET'));
    StDateArret := ' AND (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) ';
    StDateArret := StDateArret + ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
    stWhere := stWhere + StDateArret;
  End;

  // Zone de travail 1 à 4
  for i := 1 to 4 do
  begin
    if (GetControlText('PSA_TRAVAILN' + IntToStr(i)) <> '') and (GetControlText('PSA_TRAVAILN' + IntToStr(i)) <> '<<Tous>>') then
      stWhere := stWhere + ' AND PSA_TRAVAILN' + IntToStr(i) + ' = "' + GetControlText('PSA_TRAVAILN' + IntToStr(i)) + '"';
  end;

  // Code statistique
  if (GetControlText('PSA_CODESTAT') <> '') and (GetControlText('PSA_CODESTAT') <> '<<Tous>>') then
    stWhere := stWhere + ' AND PSA_CODESTAT = "' + GetControlText('PSA_CODESTAT') + '"';

  // Dates application
  if GetControlText('PHD_DATEAPPLIC') <> '' then
  begin
    DateApplic := StrtoDate(GetControlText('PHD_DATEAPPLIC'));
    stWhere := stWhere + ' AND PHD_DATEAPPLIC>="' + UsDateTime(DateApplic) + '"';
  end;
  if GetControlText('PHD_DATEAPPLIC_') <> '' then
  begin
    DateApplic := StrtoDate(GetControlText('PHD_DATEAPPLIC_'));
    stWhere := stWhere + ' AND PHD_DATEAPPLIC<="' + UsDateTime(DateApplic) + '"';
  end;

  //DEB PT1
  St := SQLConf('SALARIES');
  if St <> '' then
    St := ' AND ' + St;
  stWhere := stWhere + St;
  //FIN PT1

  SetControlText('XX_WHERE',stWhere);
  SetControlText('XX_ORDERBY','PHD_SALARIE,PHD_PGINFOSMODIF,PHD_DATEAPPLIC');
end;

procedure TOF_PGMULDETAILHISTO.OnClickSalarieSortie(Sender: TObject);
begin
  SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
  SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;

procedure TOF_PGMULDETAILHISTO.OnLoad;
begin
  inherited;
  ActiveWhere(nil);
end;


procedure TOF_PGMULDETAILHISTO.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_PGMULDETAILHISTO.PGRechSal(FF: TForm);
var i : Integer;
LeControl : TControl;
begin
  for i := 0 to FF.ComponentCount - 1 do
  begin
    LeControl := TControl(FF.Components[i]);
    if LeControl is THLabel then continue;
    if (LeControl is THEdit) AND ((pos('SALARIE', LeControl.Name) > 0) OR (pos('SALARIE_', LeControl.Name) > 0)) then
      begin
      THEdit(LeControl).OnExit := ExitEdit;
      end
      else Continue;
  end;
end;

Initialization
  registerclasses ( [ TOF_PGMULDETAILHISTO ] ) ;
end.

