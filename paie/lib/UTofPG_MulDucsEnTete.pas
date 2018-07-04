{***********UNITE*************************************************
Auteur  ...... : MF
Créé le ...... : 16/05/2001
Modifié le ... : 24/07/2001
Description .. : Multicritère de sélection des DUCS en vue de Consulation/
Suite ........ : modification/édition
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
{
 PT1 : 08/10/2001 - V536 - MF : affichage du libellé de l'organisme
                                au lieu de son code. (nouvelle liste)
 PT2 : 19/12/2001 : V571  MF :
                             1- Correction contrôle de date, résout PB qd date
                                incorrecte. Le sablier restait affiché.
                             2- on réactive affichage du libellé de l'organisme
                                au lieu de son code. (nouvelle liste et nlle vue)
 PT3 : 13/02/2003 : V595 MF  :
                             1- Correction fct ActiveWhere suite à modification
                             de la liste PGDUCSENTETE
                             Résout pb de "Nom de champ incorrect" rencontré par
                             ASSU 2000
 PT4 : 15/09/2003 :  V_421 MF  Correction FQ 10788 : Traitement des raccourcis
                               clavier (Dates de période)
 PT5 : 18/09/2003 :  V_421 MF  PgiError remplace PgiBox

 PT6 : 02/02/2007 : V_80 FC Suivant un paramètre société, contrer l'habilitation qui est faite automatique
                            en effaçant le contenu du champ XX_WHERE2
}
unit UTofPG_MulDucsEnTete;

interface
uses
  Classes, sysutils, HTB97, hmsgbox, HCtrls, Hqry, UTOF, ParamDat,
  ParamSoc, //PT6
//unused        vierge,
  Hent1,
{$IFDEF EAGLCLIENT}
  eMul;
//unused        UTOB,
//unused        MaineAgl;
{$ELSE}
  Mul;
//unused        HDB;
//unused        FE_Main;
{$ENDIF}
type
  TOF_PGMULDUCSENTETE = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
    procedure OnUpdate; override;
  private
    WW: THEdit;
    DateDeb, DateFin: THEdit;
    Q_Mul: THQuery;
    procedure ActiveWhere(Okok: Boolean);
    procedure DateElipsisclick(Sender: TObject);
    procedure Change(Sender: TObject); // PT2-1
    procedure ChercheClick(Sender: TObject); // PT2-1
  end;

implementation

uses PgOutils, PgOutils2;

procedure TOF_PGMULDUCSENTETE.OnArgument(Arguments: string);
var
  DebPer, FinPer, ExerPerEncours, MoisE, AnneeE: string;
  DebExer, FinExer, DebTrim, FinTrim, DebSem, FinSem: TDateTime;
  Cherche: TToolbarButton97; // PT2-1
begin
  inherited;

  Q_Mul := THQuery(Ecran.FindComponent('Q'));


  WW := THEdit(GetControl('XX_WHERE'));
// d PT4
  DateDeb := ThEdit(getcontrol('XX_VARIABLED'));
  DateFin := ThEdit(getcontrol('XX_VARIABLED_'));
// d PT4
  if (DateDeb <> nil) and (DateFin <> nil) then
  begin
    DateDeb.OnElipsisClick := DateElipsisclick;
    DateDeb.OnExit := Change; // PT2-1
    DateFin.OnElipsisClick := DateElipsisclick;
    DateFin.OnExit := Change; // PT2-1
  end;
// Date par défaut : La date proposée est le trimestre en cours si la période
// en cours correspond à une fin de trimestre, sinon c'est le mois en cours.
  if RendExerSocialEnCours(MoisE, AnneeE, ExerPerEncours, DebExer, FinExer) = True then
  begin
    RendPeriodeEnCours(ExerPerEnCours, DebPer, FinPer);
    RendTrimestreEnCours(StrToDate(DebPer), DebExer, FinExer, DebTrim, FinTrim, DebSem, FinSem);
    if FindeMois(StrToDate(DebPer)) <> FinTrim then
      // mois en cours
    begin
      if DateDeb <> nil then DateDeb.text := DateToStr(StrToDate(DebPer));
      if DateFin <> nil then DateFin.text := DateToStr(FindeMois(StrToDate(DebPer)));
    end
    else
      // trimestre en cours
    begin
      if DateDeb <> nil then DateDeb.text := DateToStr(DebTrim);
      if DateFin <> nil then DateFin.text := DateToStr(FinTrim);
    end;
  end;

// début PT2-1
  Cherche := TToolbarButton97(GetControl('BCHERCHE'));
  if Cherche <> nil then
  begin
    Cherche.OnClick := ChercheClick;
  end;
// fin PT2-1
end;

procedure TOF_PGMULDUCSENTETE.OnLoad;
var
  Okok: Boolean;
begin
  inherited;
  Okok := TRUE;
  ActiveWhere(Okok);
end;

procedure TOF_PGMULDUCSENTETE.OnUpdate;
begin
  ;
end;

procedure TOF_PGMULDUCSENTETE.DateElipsisclick(Sender: TObject);
var key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOF_PGMULDUCSENTETE.ActiveWhere(Okok: Boolean);
var
  St: string;
  Where2:THEdit;//PT6
begin
  WW.Text := '';
  st := '';

  St := '(PDU_DATEDEBUT >="' + UsDateTime(StrToDate(DateDeb.text)) + '") AND (PDU_DATEDEBUT <="' + UsDateTime(StrToDate(DateFin.text)) +
    '") AND (PDU_DATEFIN >="' + UsDateTime(StrToDate(DateDeb.text)) + '") AND (PDU_DATEFIN <= "' + UsDateTime(StrToDate(DateFin.text)) + '")'; //PT3
//PT3    ' AND POG_ETABLISSEMENT=PDU_ETABLISSEMENT ';     // PT2-2
  if St <> '' then WW.Text := st;
  if Q_Mul <> nil then
  begin
// PT1
// Q_Mul.Liste  := 'PGMULDUCS';
    TFMul(Ecran).SetDBListe('PGDUCSENTETE');
  end;

  //DEB PT6
  if GetParamSocSecur('SO_PGDRTVISUETAB',True) then
  begin
    Where2 := THEdit(GetControl('XX_WHERE2'));
    if Where2 <> nil then SetControlText('XX_WHERE2', '');
  end;
  //FIN PT6
end;

// debut PT2-1

procedure TOF_PGMULDUCSENTETE.Change(Sender: TObject);
begin
  if not IsValidDate(GetControlText('XX_VARIABLED')) then //  PT4
  begin //Pour générer message erreur si date erronnée
    PGIError('La date de début est erronée.', Ecran.caption); // PT5
    SetControlText('XX_VARIABLED', DatetoStr(Date)); //  PT4
  end;
  if not IsValidDate(GetControlText('XX_VARIABLED_')) then //  PT4
  begin //Pour générer message erreur si date erronnée
    PGIError('La date de fin est erronée.', Ecran.caption); // PT5
    SetControlText('XX_VARIABLED_', DatetoStr(Date)); //  PT4
  end;
end;

procedure TOF_PGMULDUCSENTETE.ChercheClick(Sender: TObject);
begin
  if not IsValidDate(GetControlText('XX_VARIABLED')) then //  PT4
  begin //Pour générer message erreur si date erronnée
    PGIError('La date de début est erronée.', Ecran.caption); // PT5
    SetControlText('XX_VARIABLED_', DatetoStr(Date)); //  PT4
  end;
  if not IsValidDate(GetControlText('XX_VARIABLED_')) then //  PT4
  begin //Pour générer message erreur si date erronnée
    PGIError('La date de fin est erronée.', Ecran.caption); // PT5
    SetControlText('XX_VARIABLED_', DatetoStr(Date)); //  PT4
  end;
  TFMul(Ecran).BChercheClick(nil);
end;
// fin PT2-1

initialization
  registerclasses([TOF_PGMULDUCSENTETE]);
end.

