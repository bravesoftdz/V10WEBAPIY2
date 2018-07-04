{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 02/08/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : REASSORTMUL ()
Mots clefs ... : TOF;REASSORTMUL
*****************************************************************}
unit REASSORTMUL_TOF;

interface

uses StdCtrls, Controls, Classes, M3FP,
  {$IFDEF EAGLCLIENT}
  eMul, UTOB,
  {$ELSE}
  Mul, db, dbtables,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox,
  UTOF, Reassort, EntGC, AglInitGC;

type
  TOF_REASSORTMUL = class(TOF)
    procedure OnArgument(S: string); override;
  end;

  TOF_REASSORTLIGNEMUL = class(TOF)
    procedure OnArgument(S: string); override;
  end;

implementation

procedure TOF_REASSORTMUL.OnArgument(S: string);
var Critere: string;
  F: TFMul;
begin
  inherited;
  F := TFMul(Ecran);
  repeat
    Critere := uppercase(Trim(ReadTokenSt(S)));
    if Critere = 'CONSULTATION' then
    begin
      F.Caption := TraduireMemoire('Consultation d''une commande de réassort');
      SetControlProperty('GP_NATUREPIECEG', 'Enabled', False);
    end
    else if Critere = 'MODIFICATION' then F.Caption := TraduireMemoire('Modification d''une commande de réassort');
  until Critere = '';
  UpdateCaption(Ecran);
end;

procedure AppelReassort(Parms: array of variant; nb: integer);
var Critere, Arguments, ChampMul, ValMul: string;
  F: TFMul;
  Latof: TOF;
  CleDoc: R_CleDoc;
  x: integer;
  Action: TActionFiche;
begin
  Action := taConsult;
  F := TFMul(Integer(Parms[0]));
  if (F is TFMul) then Latof := TFMul(F).Latof else exit;
  if (Latof is TOF_REASSORTMUL) or (Latof is TOF_REASSORTLIGNEMUL) then
  begin
    Arguments := string(Parms[1]);
    repeat
      Critere := UpperCase(Trim(ReadTokenSt(Arguments)));
      if Critere <> '' then
      begin
        x := pos('=', Critere);
        if x <> 0 then
        begin
          ChampMul := copy(Critere, 1, x - 1);
          ValMul := copy(Critere, x + 1, length(Critere));
        end;
        if ChampMul = 'NATUREPIECE' then CleDoc.NaturePiece := ValMul;
        if ChampMul = 'DATEPIECE' then CleDoc.DatePiece := StrToDate(ValMul);
        if ChampMul = 'SOUCHE' then CleDoc.Souche := ValMul;
        if ChampMul = 'NUMEROPIECE' then CleDoc.NumeroPiece := StrToInt(ValMul);
        if ChampMul = 'INDICE' then CleDoc.Indice := StrToInt(ValMul);
        if ChampMul = 'ACTION' then
        begin
          if ValMul = 'CONSULTATION' then Action := taConsult;
          if ValMul = 'MODIFICATION' then Action := taModif;
        end;
      end;
    until Critere = '';
    // modif 01/08/2001
    // if (Action=taModif) and (CleDoc.NaturePiece='TRE') then Action:=taConsult;
    SaisieCommandeReassort(CleDoc, Action);
  end;
end;

{==============================================================================================}
{=============================== TOF_RESSORTLIGNEMUL ==========================================}
{==============================================================================================}

procedure TOF_REASSORTLIGNEMUL.OnArgument(S: string);
var Critere: string;
  F: TFMul;
begin
  inherited;
  F := TFMul(Ecran);
  repeat
    Critere := uppercase(Trim(ReadTokenSt(S)));
    if Critere = 'CONSULTATION' then
    begin
      F.Caption := TraduireMemoire('Consultation d''une commande de réassort');
      SetControlProperty('GP_NATUREPIECEG', 'Enabled', False);
    end
    else if Critere = 'MODIFICATION' then F.Caption := TraduireMemoire('Modification d''une commande de réassort');
  until Critere = '';
  UpdateCaption(Ecran);
  if not VH_GC.GCMultiDepots then
  begin
    SetControlVisible('TGL_DEPOT', False);
    SetControlVisible('GL_DEPOT', False);
  end;
end;

procedure TOFREASSORTLIGNE_MUL_RechArt(Parms: array of variant; nb: integer);
var F: TForm;
  G_Article: THCritMaskEdit;
  QArt: TQuery;
  ChampArticle, NomDuChamp: string;
begin
  F := TForm(Longint(Parms[0]));
  NomDuChamp := Parms[2];
  ChampArticle := Parms[1];
  G_Article := THCritMaskEdit(F.FindComponent(ChampArticle));
  DispatchRecherche(G_Article, 1, '',
    NomDuChamp + '=' + Trim(Copy(G_Article.Text, 1, 18)) + ';GA_TYPEARTICLE=MAR', '');
  if G_Article.Text <> '' then
  begin
    THEdit(F.FindComponent('ARTICLE')).Text := G_Article.Text;
    QArt := OpenSQL('SELECT GA_CODEARTICLE FROM ARTICLE WHERE GA_ARTICLE="'
      + G_Article.Text + '"', True);
    if not QArt.EOF then
      THEdit(F.FindComponent(ChampArticle)).Text := QArt.FindField('GA_CODEARTICLE').AsString;
    Ferme(QArt);
  end;
end;

initialization
  registerclasses([TOF_REASSORTMUL]);
  registerclasses([TOF_REASSORTLIGNEMUL]);
  RegisterAglProc('AppelReassort', True, 5, AppelReassort);
  RegisterAglProc('RechArtReassortLigne_Mul', True, 1, TOFREASSORTLIGNE_MUL_RechArt);
end.
