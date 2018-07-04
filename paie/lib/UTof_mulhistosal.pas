{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. : Unit de gestion du multi critère de l'historique des
Suite ........ : évènements salariés
Mots clefs ... : PAIE;SALARIE
*****************************************************************}
{
PT1 25/03/2002 V571 PH Modif des champs historiques salarie
PT2 27/04/2006 V_65 JL FQ 12518 Supp affichage salarié confidentiel (sauf pour eManager)
PT3 26/01/2007 V_80 FC Mise en place du filtage habilitation pour les lookuplist
                       pour les critères code salarié uniquement
PT4 28/06/2007 V_72 JL FQ 14459 Gestion accès depuis fiche salarié
PT5 21/12/2007 V_81 FC FQ 14996 Concept accessibilité fiche salarié
}
unit UTof_MulHistoSal;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, DBGrids,Fe_Main,
{$ELSE}
  MaineAgl,
{$ENDIF}
  Grids, HCtrls, HEnt1, HMsgBox, UTOF, UTOB, UTOM, Vierge, P5Util, P5Def, AGLInit,
  LookUp, EntPaie, PgOutils, PgOutils2,Paramdat;

type
  TOF_MulHistoSal = class(TOF)
  private
    WW: tHedit;
    Filtr: string;
  public
    procedure OnArgument(Arguments: string); override;
  private
    Action:String;
    procedure ActiveWhere(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure OnLoad; override;
    procedure ExitEdit(Sender: TObject);
    procedure DateExit(Sender: TObject);
    procedure ClickSal(Sender: TObject);
    procedure DoubleCLickListe(Sender: TObject);   //PT5
  end;

implementation

procedure TOF_MulHistoSal.OnArgument(Arguments: string);
var
  LaDate: THEdit;
  Defaut : ThEdit;
  st: string;
  Btn:TToolbarButton97;
{$IFNDEF EAGLCLIENT}
  FListe : THDBGrid;
{$ELSE}
  FListe : THGrid;
{$ENDIF}
begin
  inherited;
  if Arguments <> '' then
  begin
    SetControlVisible('PSA_LIBELLE', FALSE);
    SetControlVisible('TPSA_LIBELLE', FALSE);
    st := Trim(Arguments);
    If Copy(St,1,8) = 'FICHESAL' then St := Copy(St,9,Length(St)) //PT4
    else SetControlText('QUOI', 'S');
    if st <> 'P' then
    SetControlText('PHS_SALARIE', ReadTokenSt(st));
    Filtr := ReadTokenSt(st);   //PT5
    Action:='';//PT5
    if St <> '' then
      Action := ReadTokenSt(St);  //PT5
    SetControlVisible('BINSERT', FALSE);
    if Filtr = 'S' then SetControlEnabled('PHS_SALARIE', FALSE);
    SetControlVisible ('BPARAMLISTE', FALSE);
    SetControlEnabled ('BFILTRE', FALSE);
    SetControlEnabled ('FFILTRES', FALSE);
  end;

  //DEB PT5
{$IFNDEF EAGLCLIENT}
  FListe := THDBGrid(GetControl('FListe'));
{$ELSE}
  FListe := THGrid(GetControl('FListe'));
{$ENDIF}
  if FListe<>nil then
    Fliste.OnDblClick:=DoubleCLickListe;
  //FIN PT5

  WW := THEdit(GetControl('XX_WHERE'));

  LaDate := THEdit(GetControl('DATEAPPLICATION'));
  if LaDate <> nil then LaDate.OnElipsisClick := DateElipsisclick;
  if Ladate <> nil then LaDate.OnExit := Dateexit;

  LaDate := THEdit(GetControl('DATEEVENEMENT'));
  if LaDate <> nil then LaDate.OnElipsisClick := DateElipsisclick;
  if Ladate <> nil then LaDate.OnExit := Dateexit;

  Defaut := ThEdit(getcontrol('PHS_SALARIE'));
  if Defaut <> nil then
  begin
    Defaut.OnExit := ExitEdit;
    Defaut.OnElipsisClick  := ClickSal ;
  end;
end;

procedure TOF_MulHistoSal.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then //AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;


procedure TOF_MulHistoSal.ActiveWhere(Sender: TObject);
var
  LaDate: ThEdit;
  StConf : String;
begin
  if WW = nil then exit;
  WW.text := '';
  LADate := THEdit(GetControl('DATEAPPLICATION'));
  if LADate <> nil then
  begin
    if LaDate.text <> '' then
      // PT1 25/03/2002 V571 PH Modif des champs historiques salarie
      WW.text := ' PHS_DATEAPPLIC >="' + usdatetime(strtodate(LaDate.text)) + '"';
  end;
  LaDate := THEdit(GetControl('DATEEVENEMENT'));
  if LaDate <> nil then
  begin
    if LaDate.text <> '' then WW.text := WW.text + ' AND ';
    WW.text := WW.Text + ' PHS_DATEEVENEMENT <="' + usdatetime(strtodate(LaDate.text)) + '"';
  end;
  if Filtr = 'P' then
  begin
    if not ConsultP then
    begin
       if WW.Text <> '' then WW.Text := WW.Text + ' AND ';
       WW.Text := WW.Text + ' EXISTS (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_SALARIE=PSA_SALARIE AND PSE_RESPONSVAR="' + LeSalarie + '")';
    end
    else
    begin
      if WW.Text <> '' then WW.Text := WW.Text + ' AND ';
      WW.Text := WW.Text + ' PSE_RESPONSABS = "' + LeSalarie + '"';
    end;
  end;
  if Filtr <> '' then
  begin
    if WW.Text <> '' then WW.Text := WW.Text + ' AND ';
    WW.Text := WW.Text + ' (PHS_BQUALIFICATION="X" OR PHS_BLIBELLEEMPLOI="X" OR PHS_BSALAIREMOIS1="X"' +
      ' OR PHS_BSALAIREMOIS2="X" OR PHS_BSALAIREMOIS3="X" OR PHS_BSALAIREMOIS4="X" OR PHS_BSALAIREANN1="X" ' +
      ' OR PHS_BHORAIREMOIS="X")';
  end;
  //DEBUT PT2
  {$IFNDEF EMANAGER}
  StConf := SQLConf('SALARIES');
  If StConf <> '' then WW.Text := WW.Text + ' AND '+StConf;
  {$ENDIF}
  //FIN PT2
end;

procedure TOF_MulHistoSal.DateExit(Sender: TObject);
var
  Ladate: thedit;
begin
  LaDate := THEdit(GetControl('DATEAPPLICATION'));
  if LADate <> nil then if not IsValidDate(LaDate.Text) then
      if LaDate.text <> '' then
      begin
        PGIBox('La date du n''est pas valide', 'Historique salarié');
        LaDate.SetFocus;
        exit;
      end;
  LADate := THEdit(GetControl('DATEEVENEMENT'));
  if LaDate <> nil then if not IsValidDate(LADate.Text) then
      if laDate.text <> '' then
      begin
        PGIBox('La date au n''est pas valide', 'Historique salarié');
        LaDate.SetFocus;
        exit;
      end;
end;

procedure TOF_MulHistoSal.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_MulHistoSal.OnLoad;
begin
  inherited;
  ActiveWhere(nil);
end;
procedure TOF_MulHistoSal.ClickSal(Sender: TObject);
var LeWhere : String ;
Defaut : THEdit ;
begin
   Defaut := ThEdit(getcontrol('PHS_SALARIE'));
  {$IFDEF EAGLCLIENT}
  if Filtr = 'P' then
  begin
//    LeWhere := 'SELECT PSA_SALARIE FROM SALARIES LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE WHERE ' ;
    if not ConsultP then
    begin
       LeWhere := LeWhere + ' EXISTS (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_SALARIE=PSA_SALARIE AND PSE_RESPONSVAR="' + LeSalarie + '")';
    end
    else
    begin
      LeWhere :=  LeWhere + ' PSE_RESPONSABS = "' + LeSalarie + '"';
    end;
  end;
{$ENDIF}
  LeWhere := RecupClauseHabilitationLookupList(LeWhere);  //PT3
  LookupList(Defaut,'Salariés' ,'SALARIES','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM',leWhere,'PSA_SALARIE',TRUE,-1) ;
end;

//DEB PT5
procedure TOF_MulHistoSal.DoubleCLickListe(Sender: TObject);
begin
  if GetField ('PHS_SALARIE') <> '' then
    if (not JaiLeDroitTag(200078)) AND (Action='ACTION=CONSULTATION') then
      AGLLanceFiche ('PAY','HISTOSAL', '',GetField('PHS_SALARIE')+';'+DateToStr(GetField('PHS_DATEEVENEMENT'))+
       ';'+DateToStr(GetField('PHS_DATEAPPLIC'))+';'+IntToStr(GetField('PHS_ORDRE')), GetField('PHS_SALARIE')+';CONSULTATION')
    else
      AGLLanceFiche ('PAY','HISTOSAL', '',GetField('PHS_SALARIE')+';'+DateToStr(GetField('PHS_DATEEVENEMENT'))+
       ';'+DateToStr(GetField('PHS_DATEAPPLIC'))+';'+IntToStr(GetField('PHS_ORDRE')), GetField('PHS_SALARIE')+';MODIFICATION');
end;
//FIN PT5

initialization
  registerclasses([TOF_MulHistoSal]);
end.

