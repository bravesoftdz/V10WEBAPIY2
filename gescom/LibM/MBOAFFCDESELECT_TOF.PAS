{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 21/11/2002
Modifié le ... : 22/11/2002
Description .. : Source TOF de la FICHE : MBO/AFFCDESELECT
Mots clefs ... : TOF;MBOAFFCDESELECT;AFFCDE
*****************************************************************}
unit MBOAFFCDESELECT_TOF;

interface

uses
  Controls, Classes, stdctrls, ComCtrls, sysutils, Forms, Windows,
  {$IFDEF EAGLCLIENT}
  eMul, Maineagl,
  {$ELSE}
  Mul, dbtables, FE_Main,
  {$ENDIF}
  UTOF, UTOB, HCtrls, HEnt1, HTB97, Hqry, HMsgBox;

type
  TOF_MBOAFFCDESELECT = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  private
    NaturePiece: string;
    procedure BrancheOnChange(NomChamp: string; Fct: TNotifyEvent);
    procedure BrancheBouton(NomBouton: string; Fct: TNotifyEvent);
    procedure MakeWhereUneLigne(QQ: THQuery; var sWhere: string);
    function MakeWhere: string;
    procedure ClickBOuvrir(Sender: TObject);
    procedure ChangeCodeAff(Sender: TObject);
    procedure ChangeCBLANCERESA(Sender: TObject);
    procedure MajChampsLibresPiece;
  end;

implementation

uses
  UtilAffCde, UtilGC, UtilRT, EntGC;

const
  NATURECDECLI = 'CC'; // nature de pièce commande client

  {***********A.G.L.Privé.*****************************************
  Auteur  ...... : N. ACHINO
  Créé le ...... : 26/11/2002
  Modifié le ... : 03/12/2002
  Description .. : Branche une fonction sur le changement d'un champ
  Mots clefs ... :
  *****************************************************************}

procedure TOF_MBOAFFCDESELECT.BrancheOnChange(NomChamp: string; Fct: TNotifyEvent);
var Ctrl: TControl;
begin
  Ctrl := GetControl(NomChamp);
  if Ctrl <> nil then
  begin
    if Ctrl is TEdit then TEdit(Ctrl).OnChange := Fct else
      if Ctrl is TCheckBox then TCheckBox(Ctrl).OnClick := Fct else
      if Ctrl is TComboBox then TComboBox(Ctrl).OnChange := Fct;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 03/12/2002
Description .. : Branche une fonction sur le click d'un bouton
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDESELECT.BrancheBouton(NomBouton: string; Fct: TNotifyEvent);
var Ctrl: TControl;
begin
  Ctrl := GetControl(NomBouton);
  if (Ctrl <> nil) and (Ctrl is TToolbarButton97) then
    TToolbarButton97(Ctrl).OnClick := Fct;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 02/12/2002
Description .. : Constitution de la clause WHERE pour une ligne de
Suite ........ : commande sélectionnée
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDESELECT.MakeWhereUneLigne(QQ: THQuery; var sWhere: string);
begin
  if QQ <> nil then
  begin
    if sWhere <> '' then sWhere := sWhere + ' OR';
    sWhere := sWhere
      + ' (GL_NATUREPIECEG="' + QQ.FindField('GL_NATUREPIECEG').asString + '"'
      + ' AND GL_SOUCHE="' + QQ.FindField('GL_SOUCHE').asString + '"'
      + ' AND GL_NUMERO="' + IntToStr(QQ.FindField('GL_NUMERO').asInteger) + '"'
      + ' AND GL_INDICEG="' + IntToStr(QQ.FindField('GL_INDICEG').asInteger) + '"'
      + ' AND GL_NUMLIGNE="' + QQ.FindField('GL_NUMLIGNE').asString + '") '
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Constitution de la clause WHERE pour la sélection des
Suite ........ : lignes de commandes
Mots clefs ... :
*****************************************************************}

function TOF_MBOAFFCDESELECT.MakeWhere: string;
var Nb, Ind: integer;
  FF: TFMul;
begin
  Result := '';
  if Ecran is TFMul then
  begin
    FF := TFMul(Ecran);
    if FF.FListe.AllSelected then
    begin
      // cas de sélection de toutes les lignes
      Result := FF.Q.Criteres;
      FF.bSelectAll.Click;
    end else
    begin
      // cas d'une sélection de certaines lignes
      Nb := FF.FListe.nbSelected;
      for Ind := 0 to Nb - 1 do
      begin
        FF.FListe.GotoLeBookmark(Ind);
        {$IFDEF EAGLCLIENT}
        FF.Q.TQ.Seek(FF.FListe.Row - 1);
        {$ENDIF}
        MakeWhereUneLigne(FF.Q, Result);
      end;
      FF.FListe.ClearSelected;
      FF.Q.First;
      if Result <> '' then Result := ' (' + Result + ') ';
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 02/12/2002
Description .. : Lancement de la recopie des lignes de commandes dans la
Suite ........ : table spécifique des affectations
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDESELECT.ClickBOuvrir(Sender: TObject);
var CodeAff, CodeStatut, Stg: string;
  Nb: integer;
begin
  if not (Ecran is TFMul) then Exit;
  // changement de critères de sélection
  if TFMul(Ecran).CritModified then
  begin
    PgiBox('Vous n''avez pas appliqué les critères !', Ecran.Caption);
    Exit;
  end;
  // contrôle du code affectation
  CodeAff := GetControlText('CODEAFF');
  if CodeAff = '' then
  begin
    PgiBox('Vous devez sélectionner un code affectation !', Ecran.Caption);
    SetFocusControl('CODEAFF');
    Exit;
  end;
  // vérifie si l'étape précédente a été faite
  if not VerifStatutAFFCDEENTETE(CodeAff, afcSelection) then Exit;
  // contrôle des lignes sélectionnées
  if (TFMul(Ecran).FListe.nbSelected = 0) and
    not (TFMul(Ecran).FListe.AllSelected) then
  begin
    PgiBox('Pour lancer le traitement, vous devez sélectionner au moins une ligne !', Ecran.Caption);
    if TFMul(Ecran).FListe.CanFocus then TFMul(Ecran).FListe.SetFocus;
    Exit;
  end;
  // Lancement du traitement
  EnableControls(TForm(Ecran), False, True);
  SetControlEnabled('BValider', False);
  SourisSablier;
  CodeStatut := DonneCodeStatutAFFCDE(afcSelection);
  MajStatutAFFCDEENTETE(CodeAff, CodeStatut, False);
  // Recopie des lignes de commandes
  Stg := MakeWhere;
  Nb := InsertAFFCDELIGNE(CodeAff, Stg);
  Stg := IntToStr(Nb) + TraduireMemoire(' ligne(s) de commandes traitée(s)');
  MajStatutAFFCDEENTETE(CodeAff, CodeStatut, True);
  // Ecriture en base de registres du code affectation utilisé
  SaveSynRegKey('CodeAffCde', CodeAff, True);
  SourisNormale;
  EnableControls(TForm(Ecran), True, True);
  // Lancement de la réservation à l'issue de la sélection
  if (Nb > 0) and (GetCheckBoxState('CBLANCERESA') = cbChecked) then
    AGLLanceFiche('MBO', 'AFFCDELANCE', '', '', 'RESERVATION;LANCEAUTO;CODEAFF=' + CodeAff)
  else
    PGIInfo(Stg, Ecran.Caption);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Changement de code affectation
Suite ........ : lignes de commandes
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDESELECT.ChangeCodeAff(Sender: TObject);
var CodeAff, Stg: string;
  QQ: TQuery;
begin
  CodeAff := GetControlText('CODEAFF');
  if CodeAff <> '' then
  begin
    Stg := 'SELECT GEA_DEPOT FROM AFFCDEENTETE WHERE GEA_CODEAFF="' + CodeAff + '"';
    QQ := OpenSQl(Stg, True);
    if QQ.Eof then
      Stg := ''
    else
      Stg := QQ.FindField('GEA_DEPOT').asString;
    Ferme(QQ);
    SetControlText('GL_DEPOT', Stg);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 05/02/2003
Description .. : Changement l'indicateur "Lancer la réservation à l'issue de
Suite ........ : la sélection" pour éviter de modifier l'indicateur de
Suite ........ : modification d'un critère.
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDESELECT.ChangeCBLANCERESA(Sender: TObject);
begin
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Paramétrage des tables libres des pièces
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDESELECT.MajChampsLibresPiece;
var Ind: integer;
  Plus, NomChamp: string;
begin
  for Ind := 1 to 3 do
  begin
    NomChamp := 'GP_LIBREPIECE' + InttoStr(Ind);
    Plus := GetInfoParPiece(NaturePiece, 'GPP_PIECETABLE' + InttoStr(Ind));
    if Plus <> '' then
    begin
      SetControlText('T' + NomChamp, RechDom('GCTYPESTATPIECE', Plus, False));
      SetControlProperty(NomChamp, 'Plus', Plus);
    end else
    begin
      SetControlVisible('T' + NomChamp, False);
      SetControlVisible(NomChamp, False);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : OnNew
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDESELECT.OnNew;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : OnDelete
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDESELECT.OnDelete;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : OnUpdate
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDESELECT.OnUpdate;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : OnLoad
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDESELECT.OnLoad;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : OnArgument
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDESELECT.OnArgument(S: string);
var Ind: integer;
  Stg: string;
begin
  inherited;
  // contraintes de sélection des lignes de commandes
  NaturePiece := NATURECDECLI; // commandes client
  SetControlText('GL_NATUREPIECEG', NaturePiece);
  SetControlChecked('GL_VIVANTE', True); // pièces vivantes
  Stg := 'GL_TYPEARTICLE="MAR" AND GL_TYPELIGNE="ART" AND GL_QTERESTE>0';
  SetControlText('XX_WHERE', Stg); // articles de type marchandise et une quantité restante 
  // code affectation
  BrancheOnChange('CODEAFF', ChangeCodeAff);
  Stg := GetSynRegKey('CodeAffCde', '', True);
  if Stg <> '' then SetControlText('CODEAFF', Stg);
  BrancheOnChange('CBLANCERESA', ChangeCBLANCERESA);
  // Paramétrage des libellés des familles, collection et stat. article
  for Ind := 1 to 3 do
    ChangeLibre2('TGL_FAMILLENIV' + InttoStr(Ind), Ecran);
  ChangeLibre2('TGL_COLLECTION', Ecran);
  MajChampsLibresArticle(TForm(Ecran));
  // Paramétrage des stat. client
  RTMajChampsLibres(TForm(Ecran));
  // Paramétrage des tables libres des pièces
  MajChampsLibresPiece;
  // Branchement de la fonction d'exécution
  BrancheBouton('BOuvrir', ClickBOuvrir);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : OnClose
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDESELECT.OnClose;
begin
  inherited;
end;

initialization
  registerclasses([TOF_MBOAFFCDESELECT]);
end.
