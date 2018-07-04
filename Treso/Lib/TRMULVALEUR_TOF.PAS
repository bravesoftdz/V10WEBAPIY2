{ Unité : Source TOF de la FICHE : TRMULVALEUR
--------------------------------------------------------------------------------------
    Version    |   Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 0.91           20/10/03   JP    Création de l'unité
 1.0.1.001.???  03/03/04   JP    Suppression à toute référence aux devises car inutiles
 6.0X.xxx.xxx   04/08/04   JP    Gestion des PopupMenus
 6.0.0.014.001  15/09/04   JP    FQ 10135 : sélection des écritures à supprimer
 6.0.0.018.001  12/10/04   JP    FQ 10177 : "Pré-création" des conditions de valeur
 7.05.001.001   23/10/06   JP    Gestion des filtres multi sociétés
--------------------------------------------------------------------------------------}
unit TRMULVALEUR_TOF ;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  StdCtrls, Controls, Classes, HTB97, 
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} mul, FE_Main, HDB,
  {$ELSE}
  eMul, MaineAGL,
  {$ENDIF}
  Forms, SysUtils, HCtrls, HEnt1, HMsgBox, uTob, UTOF, Menus;

type
  TOF_TRMULVALEUR = class (TOF)
    procedure OnArgument(S : string); override;
  private
    {$IFDEF EAGLCLIENT}
    FListe : THGrid;
    {$ELSE}
    FListe : THDBGrid;
    {$ENDIF}
    PopupMenu : TPopUpMenu; {04/08/04}

    procedure DeleteOnClick(Sender : TObject);
    procedure InsertOnClick(Sender : TObject);
    procedure DupliqOnClick(Sender : TObject);
    procedure BanqueChange (Sender : TObject);
    procedure ListeDblClick(Sender : TObject);
    procedure PreCreerClick(Sender : TObject); {12/10/04 : FQ 10177}
  end ;

procedure TRLanceFiche_MulValeur(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  TomConditionVal, Constantes, HStatus, Commun;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_MulValeur(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULVALEUR.OnArgument(S : String);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 50000156;
  TToolbarButton97(GetControl('BINSERT'   )).OnClick  := InsertOnClick;
  TToolbarButton97(GetControl('BDELETE'   )).OnClick  := DeleteOnClick;
  TToolbarButton97(GetControl('BDUPLIQUER')).OnClick  := DupliqOnClick;
  TToolbarButton97(GetControl('BPRECREER' )).OnClick  := PreCreerClick; {12/10/04 : FQ 10177}
  THValComboBox   (GetControl('TCV_BANQUE')).OnChange := BanqueChange;
  {$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FLISTE'));
  FListe.MultiSelect := True; {FQ 10135}
  {$ELSE}
  FListe := THDBGrid(GetControl('FLISTE'));
  FListe.MultiSelection := True; {FQ 10135}
  {$ENDIF}
  FListe.OnDblClick := ListeDblClick;

  PopupMenu := TPopUpMenu(GetControl('POPUPMENU')); {04/08/04}
  PopupMenu.Items[0].OnClick := InsertOnClick;
  PopupMenu.Items[1].OnClick := DupliqOnClick;
  PopupMenu.Items[2].OnClick := DeleteOnClick;
  AddMenuPop(PopupMenu, '', '');

  {23/10/06Gestion des filtres multi sociétés sur banquecp et dossier}
  THEdit(GetControl('TCV_GENERAL')).Plus := FiltreBanqueCp(THEdit(GetControl('TCV_GENERAL')).DataType, '', '');
  SetPlusBancaire(THValComboBox(GetControl('TCV_BANQUE')), 'PQ', CODECOURANTS + ';' + CODETITRES + ';');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULVALEUR.DeleteOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  n   : Integer;
begin
  if FListe.nbSelected = 0 then begin
    HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner les conditions à supprimer.;W;O;O;O;', '', '');
    Exit;
  end;

  if HShowMessage('0;' + Ecran.Caption + ';Êtes-vous sûr de vouloir supprimer les lignes sélectionnées ?;Q;YN;N;N;', '', '') = mrNo then Exit;
  {15/09/04 : FQ 10135 : On fonctionne sur les lignes sélectionnées et non plus la ligne en cours}
  for n := 0 to FListe.NbSelected - 1 do begin
    FListe.GotoLeBookmark(n);
    SQL := 'DELETE FROM CONDITIONVAL WHERE ' +
           'TCV_GENERAL = "' + VarToStr(GetField('TCV_GENERAL')) + '" AND ' +
           'TCV_CODECIB = "' + VarToStr(GetField('TCV_CODECIB')) + '" ';
    ExecuteSQL(SQL);
  end;

  {Pour raffraichir la liste}
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULVALEUR.InsertOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_ConditionVal('TR','TRCONDITIONVAL', '', '', 'ACTION=CREATION');
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULVALEUR.DupliqOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_ConditionVal('TR','TRCONDITIONVAL', '', '', 'ACTION=CREATION;' +
                                                           VarToStr(GetField('TCV_GENERAL')) + ';' +
                                                           VarToStr(GetField('TCV_CODECIB')) + ';' +
                                                           CODEDUPLICAT + ';');
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULVALEUR.ListeDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_ConditionVal('TR','TRCONDITIONVAL', '', GetField('TCV_GENERAL') + ';' +
                                                       GetField('TCV_CODECIB'), 'ACTION=MODIFICATION');
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULVALEUR.BanqueChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  SetControlText('TCV_GENERAL', '');
  SetControlText('TCV_CODECIB', '');

  if GetControlText('TCV_BANQUE') = '' then begin
    {23/10/06Gestion des filtres multi sociétés sur banquecp et dossier}
    THEdit(GetControl('TCV_GENERAL')).Plus := FiltreBanqueCp(THEdit(GetControl('TCV_GENERAL')).DataType, '', '');
    THEdit(GetControl('TCV_CODECIB')).Plus := '';
  end
  else begin
    {23/10/06Gestion des filtres multi sociétés sur banquecp et dossier}
    THEdit(GetControl('TCV_GENERAL')).Plus := FiltreBanqueCp(THEdit(GetControl('TCV_GENERAL')).DataType, '', '') + 
                                              ' AND BQ_BANQUE = "' + THValComboBox(GetControl('TCV_BANQUE')).Value + '"';

    THEdit(GetControl('TCV_CODECIB')).Plus := 'TCI_BANQUE = "' + THValComboBox(GetControl('TCV_BANQUE')).Value + '"';
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULVALEUR.PreCreerClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  tBqe  : TOB;
  tCIB  : TOB;
  tVal  : TOB;
  T     : TOB;
  n, p  : Integer;
  QQ    : TQuery;
  WhBqe : string;
  WhCib : string;
  WhVal : string;
  Nb    : Integer;
begin
  if HShowMessage('0;' + Ecran.Caption + ';Êtes-vous sûr de vouloir créer les conditions de valeur à vide ?;Q;YNC;N;C;', '', '') <> mrYes then Exit;

  {Constitution des clauses Where sur les tables BANQUECP, CONDITIONVAL et CIB}
  if GetControlText('TCV_GENERAL') <> '' then begin
    WhBqe := 'WHERE BQ_CODE = "' + GetControlText('TCV_GENERAL') + '"';
    QQ := OpenSQL('SELECT BQ_BANQUE FROM BANQUECP WHERE BQ_CODE = "' + GetControlText('TCV_GENERAL') + '"', True);
    if not QQ.EOF then
      WhCib := 'WHERE TCI_BANQUE = "' + QQ.FindField('BQ_BANQUE').AsString + '"';
    Ferme(QQ);
    WhVal := 'WHERE TCV_GENERAL = "' + GetControlText('TCV_GENERAL') + '"';
  end
  else if GetControlText('TCV_BANQUE') <> '' then begin
    WhBqe := 'WHERE BQ_BANQUE = "' + GetControlText('TCV_BANQUE') + '"';
    WhCib := 'WHERE TCI_BANQUE = "' + GetControlText('TCV_BANQUE') + '"';
    WhVal := 'WHERE TCV_BANQUE = "' + GetControlText('TCV_BANQUE') + '"';
  end
  else
    WhCib := 'WHERE TCI_BANQUE = "' + CODECIBREF + '"';

  Nb := 0;
  tBqe := TOB.Create('***', nil, -1);
  tCIB := TOB.Create('µµµ', nil, -1);
  tVal := TOB.Create('$$$', nil, -1);
  try
    tBqe.LoadDetailFromSQL('SELECT BQ_CODE, BQ_BANQUE FROM BANQUECP ' + WhBqe);
    tCIB.LoadDetailFromSQL('SELECT TCI_CODECIB FROM CIB ' + WhCib);
    tVal.LoadDetailDBFromSQL('CONDITIONVAL', 'SELECT * FROM CONDITIONVAL ' + WhVal, True);

    InitMove(tBqe.Detail.Count * tCib.Detail.Count, Ecran.Caption);
    {On boucle sur les comptes bancaires et les CIB qui sont constitutifs de l'index de CONDITIONVAL}
    for n := 0 to tBqe.Detail.Count - 1 do begin
      for p := 0 to tCib.Detail.Count - 1 do begin
        MoveCur(False);
        T := tVal.FindFirst(['TCV_GENERAL', 'TCV_CODECIB'], [tBqe.Detail[n].GetString('BQ_CODE'), tCIB.Detail[p].GetString('TCI_CODECIB')], False);
        if not Assigned(T) then begin
          T := TOB.Create('CONDITIONVAL', tVal, -1);
          T.SetString('TCV_GENERAL', tBqe.Detail[n].GetString('BQ_CODE'));
          T.SetString('TCV_BANQUE', tBqe.Detail[n].GetString('BQ_BANQUE'));
          T.SetString('TCV_CODECIB',tCIB.Detail[p].GetString('TCI_CODECIB'));
          T.SetInteger('TCV_NBJPREMIER', 0);
          T.SetInteger('TCV_NBJDEUXIEME', 0);
          T.SetString('TCV_TYPEGLISSEMENT', '0');
          T.SetString('TCV_TYPEPREMIER', '0');
          T.SetString('TCV_TYPEDEUXIEME', '0');
          T.InsertDb(nil, False);
          Inc(Nb);
        end;
      end;
    end;
    FiniMove;

    if Nb = 0 then WhCib := TraduireMemoire('Aucune condition de valeur n''a été créée.')
    else if Nb = 1 then WhCib := TraduireMemoire('Une condition de valeur a été créée.')
    else WhCib := IntToStr(Nb) + TraduireMemoire(' conditions de valeur ont été créées.');
    HShowMessage('1;' + Ecran.Caption + ';' + WhCib + ';I;O;O;O;', '', '');
    if Nb > 0 then
      TFMul(Ecran).BChercheClick(Sender);
  finally
    FreeAndNil(tBqe);
    FreeAndNil(tCIB);
    FreeAndNil(tVal);
  end;
end;

initialization
  RegisterClasses([TOF_TRMULVALEUR]);

end.
