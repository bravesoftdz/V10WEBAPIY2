{***********UNITE*************************************************
Auteur  ...... : A. CATHELINEAU
Créé le ...... : 26/11/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFFCDESAISIE ()
Mots clefs ... : TOF;AFFCDESAISIE
*****************************************************************}
unit AFFCDESAISIE_TOF;

interface

uses StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, dbtables,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, Vierge, UtilArticle;

type
  TOF_AFFCDESAISIE = class(TOF)
  private
    Article, Tiers, TypeAction, Modifiable, Modifier: string;
    QteALivrer, QteReservee, QteAffectee, QteDispo, QteStock: Double;
    NotClose: Boolean;
    function VerifStock: boolean;

  public
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  end;

implementation

procedure TOF_AFFCDESAISIE.OnNew;
begin
  inherited;
end;

procedure TOF_AFFCDESAISIE.OnDelete;
begin
  inherited;
end;

procedure TOF_AFFCDESAISIE.OnUpdate;
begin
  inherited;
  if TypeAction = 'AFFECTATION' then
  begin
    if VerifStock then
    begin
      SetFocusControl('QTEAFFECTEE');
      NotClose := True;
    end
    else
    begin
      if GetCheckBoxState('MODIFIABLE') = cbChecked then Modifiable := 'X' else Modifiable := '-';
      QteAffectee := StrToFloat(GetControlText('QTEAFFECTEE'));
      Modifier := 'X';
    end;
  end
  else
  begin
    if GetCheckBoxState('MODIFIABLE') = cbChecked then Modifiable := 'X' else Modifiable := '-';
    QteReservee := StrToFloat(GetControlText('QTERESERVEE'));
    Modifier := 'X';
  end;
end;

procedure TOF_AFFCDESAISIE.OnLoad;
begin
  inherited;
  Modifier := '-';
  if TypeAction = 'RESERVATION' then
  begin
    Ecran.Caption := 'Réservation manuelle des commandes ';
    SetControlProperty('QTEAFFECTEE', 'Visible', False);
    SetControlProperty('TQTEAFFECTEE', 'Visible', False);
    SetFocusControl('QTEAFFECTEE');
  end
  else
  begin
    Ecran.Caption := 'Affectation manuelle des commandes ';
    SetControlProperty('QTERESERVEE', 'Enabled', False);
    SetFocusControl('QTERESERVEE');
  end;
  SetControlCaption('TLARTICLE', LibelleArticleGenerique(Article));
  SetControlText('QTESTOCK', FloatToStr(QteStock));
  SetControlText('QTEDISPO', FloatToStr(QteDispo));
  SetControlText('QTEALIVRER', FloatToStr(QteALivrer));
  SetControlText('QTERESERVEE', FloatToStr(QteReservee));
  SetControlText('QTEAFFECTEE', FloatToStr(QteAffectee));
  SetControlChecked('MODIFIABLE', Boolean(Modifiable = 'X'));
end;

procedure TOF_AFFCDESAISIE.OnArgument(S: string);
var Critere, ChampMul, ValMul: string;
  x: Integer;
begin
  inherited;
  //Action:=taConsult ;
  repeat
    Critere := Trim(ReadTokenSt(S));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'QteALivrer' then QteALivrer := StrToFloat(ValMul);
        if ChampMul = 'QteReservee' then QteReservee := StrToFloat(ValMul);
        if ChampMul = 'QteAffectee' then QteAffectee := StrToFloat(ValMul);
        if ChampMul = 'QteDispo' then QteDispo := StrToFloat(ValMul);
        if ChampMul = 'QteStock' then QteStock := StrToFloat(ValMul);
        if ChampMul = 'Modifiable' then Modifiable := ValMul;
        if ChampMul = 'Article' then Article := ValMul;
        if ChampMul = 'Tiers' then Tiers := ValMul;
        if ChampMul = 'TypeAction' then TypeAction := ValMul;
      end;
    end;
  until Critere = '';
end;

procedure TOF_AFFCDESAISIE.OnClose;
var QteModifie: string;
begin
  inherited;
  if NotClose then LastError := -1 else LastError := 0;
  NotClose := False;
  QteModifie := FloatToStr(QteReservee) + ';' + FloatToStr(QteAffectee) + ';' + Modifiable + ';' + Modifier;
  TFVierge(Ecran).FRetour := QteModifie;
end;

function TOF_AFFCDESAISIE.VerifStock: Boolean;
var QteSaisie: Double;
begin
  Result := False;
  QteSaisie := StrToFloat(GetControlText('QTEAFFECTEE'));
  if QteSaisie = 0 then exit;
  if QteSaisie > QteALivrer then
  begin
    PGIInfo(TraduireMemoire('La quantité demandée est supérieur à la quantité à livrer'), Ecran.Caption);
    SetFocusControl('QTEAFFECTEE');
    Result := True;
  end
  else
    if QteSaisie > QteDispo then
  begin
    PGIInfo(TraduireMemoire('Quantité insuffisante pour cet article'), Ecran.Caption);
    SetFocusControl('QTEAFFECTEE');
    Result := True;
  end;
end;

initialization
  registerclasses([TOF_AFFCDESAISIE]);
end.
