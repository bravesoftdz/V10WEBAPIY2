{ Unité : Source TOF du Mul : CPMULBAP ()
--------------------------------------------------------------------------------------
    Version   |   Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 7.01.001.001  02/02/06   JP   Création de l'unité
 8.01.001.003  26/01/07   JP   FQ 19590 : HelpTopic de la suppression
 8.01.001.003  29/01/07   JP   FQ 19447 : Choix des boutons en suivi des rejets
 8.00.001.012  24/04/07   JP   FQ 19949 : Ajout d'un mul affichant tous les BAP
 8.00.001.018  29/05/07   JP   Mise en place des rôles à la place des groupes pour filtrer les utilisateurs
 8.00.001.022  26/06/07   JP   Mise en place d'un filtre sur les auxiliaires dans le mul tym_Tous (Demande de SIC)
 8.10.002.002  30/10/07   JP   FQ 21432 : Gestion des Glyph
--------------------------------------------------------------------------------------}
unit CPMULBAP_TOF;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAGL, eMul,
  {$ELSE}
  FE_Main, db, {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF} Mul, HDB,
  {$ENDIF}
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  sysutils, UTob, ComCtrls, HCtrls, uTOF;

const
  tym_Relance     = 'R';
  tym_Modif       = 'M';
  tym_Suivi       = 'S';
  tym_Purge       = 'P';
  tym_Suppression = 'D';
  tym_Tous        = 'T';

type
  TOF_CPMULBAP = class (TOF)
    procedure OnArgument(S : string); override;
  private
    TypeAppel : Char;

    procedure SetClauseWhere;
    procedure MajControl;
    procedure ModifBap;
    procedure RelanceBap;
    procedure PurgeBap;
  public
    {$IFDEF EAGLCLIENT}
    FListe : THGrid;
    {$ELSE}
    FListe : THDBGrid;
    {$ENDIF EAGLCLIENT}

    procedure BOuvrirClick(Sender : TObject);
    procedure BTraiteClick(Sender : TObject);
    procedure ListeDbClick(Sender : TObject);
    procedure SlctAllClick(Sender : TObject);
    procedure CasesOnClick(Sender : TObject);
    procedure ElipsisClick(Sender : TObject); {29/05/07}
    procedure TiersElipsis(Sender : TObject); {26/06/07}
    procedure BDetailOnClick(Sender : TObject);{JP 13/05/08 : FQ 22665}
  end;

procedure CpLanceFiche_MulBap(TypeMul : string);

implementation

uses
  {$IFDEF EAGLCLIENT}
  UtileAgl,
  {$ELSE}
  EdtREtat,
  {$ENDIF EAGLCLIENT}
  HEnt1, HMsgBox, Ent1, HTB97, ULibBonAPayer, CPMODIFBAP_TOF, CPMODIFSTATUTBAP_TOF,
  UProcGen, AglInit, Forms, CPGENEREBAPCREATE_TOF, CPBONSAPAYER_TOF, LookUp, ParamSoc;

{---------------------------------------------------- -----------------------------------}
procedure CpLanceFiche_MulBap(TypeMul : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('CP', 'CPMULBAP', '', '', TypeMul);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULBAP.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  TypeAppel := StrToChr(ReadTokenSt(S));
  case TypeAppel of
    tym_Relance     : TFMul(Ecran).SetDBListe('CPLISTERELANCEBAP');
    tym_Modif       : TFMul(Ecran).SetDBListe('CPLISTEMODIFBAP');
    tym_Suivi       : TFMul(Ecran).SetDBListe('CPLISTESUIVIBAP');
    tym_Purge       : TFMul(Ecran).SetDBListe('CPLISTEPURGEBAP');
    tym_Suppression : TFMul(Ecran).SetDBListe('CPSUPPRESSIONBAP');
    tym_Tous        : TFMul(Ecran).SetDBListe('CPAFFICHEALLBAP'); {JP 24/04/07 : FQ 19949}
  end;

  FListe := TFMul(Ecran).FListe;
  {27/04/06 : Teste pour compatibilité entre socref 744 et version 7.0.0.003 : A supprimer ensuite}
  if Assigned(GetControl('BOUVRIR1')) then
    TToolbarButton97(GetControl('BOUVRIR1')).OnClick := ListeDbClick;
  if Assigned(GetControl('BOUVRIR2')) then
    TToolbarButton97(GetControl('BOUVRIR2')).OnClick := BTraiteClick;

  TToolbarButton97(GetControl('BOUVRIR'   )).OnClick := BOuvrirClick;
  TToolbarButton97(GetControl('BSELECTALL')).OnClick := SlctAllClick;
  TFMul(Ecran).FListe.OnDblClick := ListeDbClick;

  SetPlusCombo(THValComboBox(GetControl('BAP_JOURNAL')) , 'J');
  {JP 29/05/07 : Mise en place des rôles}
  if GetControl('BAP_VISEUR1') is THEdit then begin
    (GetControl('BAP_VISEUR1_') as THEdit).OnElipsisClick := ElipsisClick;
    (GetControl('BAP_VISEUR2' ) as THEdit).OnElipsisClick := ElipsisClick;
    (GetControl('BAP_VISEUR2_') as THEdit).OnElipsisClick := ElipsisClick;
    (GetControl('BAP_VISEUR1' ) as THEdit).OnElipsisClick := ElipsisClick;
  end else begin
    SetPlusCombo(THValComboBox(GetControl('BAP_VISEUR1')) , 'US');
    SetPlusCombo(THValComboBox(GetControl('BAP_VISEUR1_')), 'US');
    SetPlusCombo(THValComboBox(GetControl('BAP_VISEUR2')) , 'US');
    SetPlusCombo(THValComboBox(GetControl('BAP_VISEUR2_')), 'US');
  end;

  {26/06/07 : Mise en place d'un filtre sur les auxiliaires}
  SetControlVisible('E_AUXILIAIRE' , TypeAppel = tym_Tous);
  SetControlVisible('TE_AUXILIAIRE', TypeAppel = tym_Tous);
  (GetControl('E_AUXILIAIRE') as THEdit).OnElipsisClick := TiersElipsis;

  {JP 13/05/08 : FQ 22665 : récapitulatif du bon à payer}
  (GetControl('BDETAIL') as TToolbarButton97).OnClick := BDetailOnClick;

  SetClauseWhere;
  MajControl;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULBAP.SlctAllClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Fiche : TFMul;
begin
  Fiche := TFMul(Ecran);
  {$IFDEF EAGLCLIENT}
  if not Fiche.FListe.AllSelected then begin
    if not Fiche.FetchLesTous then Exit;
  end;
  {$ENDIF}
  Fiche.bSelectAllClick(nil);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULBAP.BOuvrirClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  case TypeAppel of
    tym_Relance     : RelanceBap;
    tym_Modif       : ModifBap;
    tym_Suivi       : ModifBap;
    tym_Purge       : PurgeBap;
    tym_Suppression : PurgeBap;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULBAP.SetClauseWhere;
{---------------------------------------------------------------------------------------}
var
  ch : string;
begin
  case TypeAppel of
    tym_Relance     : ch := 'BAP_STATUTBAP  = "' + sbap_Encours + '" AND BAP_ALERTE1 = "X"';
    tym_Modif       : ch := 'BAP_STATUTBAP IN ("' + sbap_Encours + '", "' + sbap_Refuse + '")';
    tym_Suivi       : ch := 'BAP_STATUTBAP IN ("' + sbap_Analytique + '", "' + sbap_Refuse + '", "' + sbpa_bloque + '")';
    tym_Purge       : ch := 'BAP_STATUTBAP = "' + sbap_Definitif + '" AND E_VALIDE = "X"';
    {JP 13/05/08 : FQ 22689 : Ajout du préfixe B1 + Mise à jour de la liste CPSUPPRESSIONBAP}
    tym_Suppression : ch := 'BAP_STATUTBAP <> "' + sbap_Definitif + '" AND BAP_NUMEROORDRE = ' +
                            '(SELECT MAX(BAP_NUMEROORDRE) FROM CPBONSAPAYER B2 WHERE B2.BAP_JOURNAL = B1.BAP_JOURNAL ' +
                            'AND B2.BAP_EXERCICE = B1.BAP_EXERCICE AND  B2.BAP_DATECOMPTABLE = B1.BAP_DATECOMPTABLE ' +
                            'AND B2.BAP_NUMEROPIECE = B1.BAP_NUMEROPIECE)';
    tym_Tous        : ch := '';        
  end;
  SetControlText('XX_WHERE', ch);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULBAP.MajControl;
{---------------------------------------------------------------------------------------}
var
  Cap : string;
begin
  if TypeAppel in [tym_Relance, tym_Modif, tym_Suppression, tym_Purge] then
    {JP 30/10/07 : FQ 21432 : Glyph du bouton valider "VVert"}
    TFMul(Ecran).BOuvrir.GlobalIndexImage := 'Z0184_S16G1';

  if TypeAppel = tym_Relance then begin
    Cap := TraduireMemoire('Relance des pièces à viser');
    SetControlVisible('BAP_RELANCE1'  , True);
    SetControlVisible('BAP_RELANCE2'  , True);
    SetControlVisible('TBAP_STATUTBAP', False);
    SetControlVisible('BAP_STATUTBAP' , False);
    TCheckbox(GetControl('BAP_RELANCE1')).OnClick := CasesOnClick;
    TCheckbox(GetControl('BAP_RELANCE2')).OnClick := CasesOnClick;
    //THValComboBox(GetControl('BAP_STATUTBAP')).Plus := 'AND (CO_CODE = "' + sbap_Encours + '")';
    TFMul(Ecran).BOuvrir.Hint := Cap;
    Ecran.HelpContext := 7509100;
  end
  else if TypeAppel = tym_Modif then begin
    Cap := TraduireMemoire('Modification des circuits de validation');
    TFMul(Ecran).BOuvrir.Hint := Cap;
    THValComboBox(GetControl('BAP_STATUTBAP')).Plus := 'AND (CO_CODE IN ("' + sbap_Encours + '", "' + sbap_Refuse + '"))';
    Ecran.HelpContext := 7509070;
  end
  else if TypeAppel = tym_Suivi then begin
    Cap := TraduireMemoire('Suivi des rejets');
    TFMul(Ecran).BOuvrir.Hint := TraduireMemoire('Modification des circuits de validation');
    {29/01/07 : FQ 19447 : ajout d'un nouveau bouton}
    if Assigned(GetControl('BOUVRIR2')) then
      TToolbarButton97(GetControl('BOUVRIR2')).Hint := TraduireMemoire('Modification du statut de l''étape');

    {JP 29/01/07 : FQ 19447 : On met un simple VVERT
     JP 30/10/07 : FQ 21432 : Glyph du bouton valider "VVert"}
    TFMul(Ecran).BOuvrir.GlobalIndexImage := 'Z0003_S16G2';

    SetControlVisible('BBLOCNOTE', True);
    SetControlVisible('BOUVRIR1'  , True);
    SetControlVisible('BOUVRIR2'  , True); {29/01/07 : FQ 19447}
    SetControlVisible('BSELECTALL', False);
    THValComboBox(GetControl('BAP_STATUTBAP')).Plus := 'AND (CO_CODE IN ("' + sbap_Analytique + '", "' + sbap_Refuse + '", "' + sbpa_bloque + '"))';
    {$IFDEF EAGLCLIENT}
    FListe.MultiSelect := False;
    {$ELSE}
    FListe.MultiSelection := False;
    {$ENDIF EAGLCLIENT}
    Ecran.HelpContext := 7509080;
  end
  else if TypeAppel = tym_Suppression then begin
    Cap := TraduireMemoire('Suppression des bons à payer');
    TFMul(Ecran).BOuvrir.Hint := Cap;
    THValComboBox(GetControl('BAP_STATUTBAP')).Plus := 'AND (CO_CODE <> "' + sbap_Definitif + '")';
    Ecran.HelpContext := 7509105; {FQ 19590 : 26/01/07}
  end
  else if TypeAppel = tym_Purge then begin
    Cap := TraduireMemoire('Purge des bons à payer');
    TFMul(Ecran).BOuvrir.Hint := Cap;
    THValComboBox(GetControl('BAP_STATUTBAP')).Plus := 'AND (CO_CODE = "' + sbap_Definitif + '")';
    Ecran.HelpContext := 7509110;
  end
  {24/04/07 : FQ 19949 : Nouveau menu}
  else if TypeAppel = tym_Tous then begin
    SetControlVisible('BOUVRIR1', True);
    SetControlVisible('BOUVRIR' , False);
    SetControlVisible('BSELECTALL', False);
    {JP 13/05/08 : FQ 22665 : visualisation du bloc note}
    SetControlVisible('BBLOCNOTE', True);
    SetControlVisible('BDETAIL', True);
    
    {$IFDEF EAGLCLIENT}
    //FListe.MultiSelect := False;
    {$ELSE}
    //FListe.MultiSelection := False;
    {$ENDIF EAGLCLIENT}
    Cap := TraduireMemoire('Suivi des bons à payer');
    Ecran.HelpContext := 150;
  end;
  Ecran.caption := Cap;
  UpdateCaption(Ecran);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULBAP.ListeDbClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  B : string;
begin
  {09/01/08 : FQ 22147 : possibilité d'ouvrir la saisie en Modification en suivi des rejets tym_Suivi}
  if TypeAppel = tym_Suivi then B := 'OUI;'
                           else B := 'NON;';
  CpLanceFiche_FicheBap(VarToStr(GetField('BAP_JOURNAL')) + ',' + VarToStr(GetField('BAP_EXERCICE')) + ',' +
                        VarToStr(GetField('BAP_DATECOMPTABLE')) + ',' + VarToStr(GetField('BAP_NUMEROPIECE')) + ',;' + B);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULBAP.ModifBap;
{---------------------------------------------------------------------------------------}
var
  {$IFDEF EAGLCLIENT}
  F : THGrid;
  {$ELSE}
  F : THDBGrid;
  {$ENDIF}
  n : Integer;
  aTob : TOB;

    {----------------------------------------------------------------------}
    procedure _AddEnreg;
    {----------------------------------------------------------------------}
    var
      T : TOB;
    begin
      T := TOB.Create('µµµ', aTob, -1);
      T.AddChampSupValeur('BAP_JOURNAL'      , TFMul(Ecran).Q.FindField('BAP_JOURNAL'      ).AsString);
      T.AddChampSupValeur('BAP_EXERCICE'     , TFMul(Ecran).Q.FindField('BAP_EXERCICE'     ).AsString);
      T.AddChampSupValeur('BAP_DATECOMPTABLE', TFMul(Ecran).Q.FindField('BAP_DATECOMPTABLE').AsDateTime);
      T.AddChampSupValeur('BAP_NUMEROPIECE'  , TFMul(Ecran).Q.FindField('BAP_NUMEROPIECE'  ).AsInteger);
      T.AddChampSupValeur('BAP_NUMEROORDRE'  , TFMul(Ecran).Q.FindField('BAP_NUMEROORDRE'  ).AsInteger);
      T.AddChampSupValeur('BAP_NBJOUR'       , TFMul(Ecran).Q.FindField('BAP_NBJOUR'       ).AsInteger);
    end;

begin
  {$IFDEF EAGLCLIENT}
  F := THGrid(TFMul(Ecran).FListe);
  {$ELSE}
  F := THDBGrid(TFMul(Ecran).FListe);
  {$ENDIF}

  aTob := TOB.Create('µµµ', nil, -1);
  try
    if TypeAppel <> tym_Suivi then begin
      {Aucune sélection, on sort}
      if (F.NbSelected = 0) and not F.AllSelected then begin
        HShowMessage('0;' + Ecran.Caption + ';Aucun élément n''est sélectionné.;W;O;O;O;', '', '');
        Exit;
      end;

      {$IFNDEF EAGLCLIENT}
      TFMul(Ecran).Q.First;
      if F.AllSelected then
        while not TFMul(Ecran).Q.EOF do begin
          _AddEnreg;
          TFMul(Ecran).Q.Next;
        end
      else
      {$ENDIF}

      for n := 0 to F.nbSelected - 1 do begin
        F.GotoLeBookmark(n);
        {$IFDEF EAGLCLIENT}
        TFMul(Ecran).Q.TQ.Seek(F.Row - 1);
        {$ENDIF}
        _AddEnreg;
      end;
    end
    else begin
      if VarToStr(GetField('BAP_STATUTBAP')) = sbap_Refuse then
        _AddEnreg
      else begin
        HShowMessage('0;' + Ecran.Caption + ';Il n''est possible de modifier que les bons réfusés.;W;O;O;O;', '', '');
        Exit;
      end;
    end;

    TheTob := aTob;
    CpLanceFiche_ModifBap;
    TFMul(Ecran).BCherche.Click;
  finally
    if Assigned(aTob) then FreeAndNil(aTob);
  end;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULBAP.PurgeBap;
{---------------------------------------------------------------------------------------}
var
  {$IFDEF EAGLCLIENT}
  F : THGrid;
  {$ELSE}
  F : THDBGrid;
  {$ENDIF}
  n : Integer;

    {----------------------------------------------------------------------}
    procedure _DeleteEnreg;
    {----------------------------------------------------------------------}
    begin
      ExecuteSQL('DELETE FROM CPBONSAPAYER WHERE BAP_JOURNAL = "' + TFMul(Ecran).Q.FindField('BAP_JOURNAL').AsString +
                 '" AND BAP_EXERCICE = "' + TFMul(Ecran).Q.FindField('BAP_EXERCICE').AsString +
                 '" AND BAP_NUMEROPIECE = ' + TFMul(Ecran).Q.FindField('BAP_NUMEROPIECE').AsString);
    end;

begin
  {$IFDEF EAGLCLIENT}
  F := THGrid(TFMul(Ecran).FListe);
  {$ELSE}
  F := THDBGrid(TFMul(Ecran).FListe);
  {$ENDIF}

  {Aucune sélection, on sort}
  if (F.NbSelected = 0) and not F.AllSelected then begin
    HShowMessage('0;' + Ecran.Caption + ';Aucun élément n''est sélectionné.;W;O;O;O;', '', '');
    Exit;
  end
  else if HShowMessage('0;' + Ecran.Caption + ';Êtes vous certain de vouloir supprimer les bons à payer sélectionnés ?;Q;YN;N;N;', '', '') <> mrYes then
    Exit;

  {$IFNDEF EAGLCLIENT}
  TFMul(Ecran).Q.First;
  if F.AllSelected then
    while not TFMul(Ecran).Q.EOF do begin
      _DeleteEnreg;
      TFMul(Ecran).Q.Next;
    end
  else
  {$ENDIF}

  for n := 0 to F.nbSelected - 1 do begin
    F.GotoLeBookmark(n);
    {$IFDEF EAGLCLIENT}
    TFMul(Ecran).Q.TQ.Seek(F.Row - 1);
    {$ENDIF}
    _DeleteEnreg;
  end;

  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULBAP.RelanceBap;
{---------------------------------------------------------------------------------------}
var
  {$IFDEF EAGLCLIENT}
  F : THGrid;
  {$ELSE}
  F : THDBGrid;
  {$ENDIF}
  n : Integer;
  ListeErr : TList;

    {----------------------------------------------------------------------}
    procedure _RelanceEnreg;
    {----------------------------------------------------------------------}
    var
      Obj  : TObjMail;
      Clef : TClefPiece;
      PRec : PClefPiece;
      s    : string;
    begin
      Clef.E_JOURNAL       := TFMul(Ecran).Q.FindField('BAP_JOURNAL'      ).AsString;
      Clef.E_EXERCICE      := TFMul(Ecran).Q.FindField('BAP_EXERCICE'     ).AsString;
      Clef.E_DATECOMPTABLE := TFMul(Ecran).Q.FindField('BAP_DATECOMPTABLE').AsDateTime;
      Clef.E_NUMEROPIECE   := TFMul(Ecran).Q.FindField('BAP_NUMEROPIECE'  ).AsInteger;
      Clef.E_NUMLIGNE      := TFMul(Ecran).Q.FindField('BAP_NUMEROORDRE'  ).AsInteger;
      Obj := TObjMail.Create(Clef);
      try
        if TFMul(Ecran).Q.FindField('BAP_RELANCE1').AsString = '-' then s := Obj.CreerMail(nma_Relance1)
                                                                   else s := Obj.CreerMail(nma_Relance2);
        if s <> '' then begin
          System.New(PRec);
          PRec^.E_JOURNAL := Clef.E_JOURNAL;
          PRec^.E_DATECOMPTABLE := Clef.E_DATECOMPTABLE;
          PRec^.E_NUMEROPIECE := Clef.E_NUMEROPIECE;
          PRec^.E_EXERCICE := Clef.E_EXERCICE;
          PRec^.E_QUALIFPIECE := s;
          ListeErr.Add(PRec);
        end;
      finally
        if Assigned(Obj) then FreeAndNil(Obj);
      end;
    end;

begin
  {$IFDEF EAGLCLIENT}
  F := THGrid(TFMul(Ecran).FListe);
  {$ELSE}
  F := THDBGrid(TFMul(Ecran).FListe);
  {$ENDIF}

  {Aucune sélection, on sort}
  if (F.NbSelected = 0) and not F.AllSelected then begin
    HShowMessage('0;' + Ecran.Caption + ';Aucun élément n''est sélectionné.;W;O;O;O;', '', '');
    Exit;
  end
  else begin
    if HShowMessage('0;' + Ecran.Caption + ';Êtes vous sûr de vouloir relancer les bons sélectionnés ?;Q;YN;N;N;', '', '') <> mrYes then
    Exit;
  end;
  ListeErr := TList.Create;
  try
    {$IFNDEF EAGLCLIENT}
    TFMul(Ecran).Q.First;
    if F.AllSelected then
      while not TFMul(Ecran).Q.EOF do begin
        _RelanceEnreg;
        TFMul(Ecran).Q.Next;
      end
    else
    {$ENDIF}

    for n := 0 to F.nbSelected - 1 do begin
      F.GotoLeBookmark(n);
      {$IFDEF EAGLCLIENT}
      TFMul(Ecran).Q.TQ.Seek(F.Row - 1);
      {$ENDIF}
      _RelanceEnreg;
    end;

    if ListeErr.Count > 0 then begin
      TheData := ListeErr;
      TheTob  := nil;
      CpLanceFiche_GenereBapColl;

    end;
  finally
    if Assigned(ListeErr) then DisposeListe(ListeErr, True);
  end;

  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULBAP.CasesOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  BOk : Boolean;
begin
  if TCheckbox(Sender).Name = 'BAP_RELANCE1' then begin
    BOk := GetCheckBoxState('BAP_RELANCE1') = cbUnchecked;
    TCheckbox(GetControl('BAP_RELANCE2')).State := cbUnchecked;
    SetControlEnabled('BAP_RELANCE2', not BOk);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULBAP.BTraiteClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  CpLanceFiche_ModifStatutBap(VarToStr(GetField('BAP_JOURNAL')) + ',' + VarToStr(GetField('BAP_EXERCICE')) + ',' +
                              VarToStr(GetField('BAP_DATECOMPTABLE')) + ',' + VarToStr(GetField('BAP_NUMEROPIECE')) +
                              ',' + VarToStr(GetField('BAP_NUMEROORDRE')) + ',;');

  TFMul(Ecran).BCherche.Click;
end;

{JP 29/05/07 : Mise en place des rôles
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULBAP.ElipsisClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  LookUpUtilisateur(Sender);
end;

{26/06/07 : Gestion des auxiliaires quand TypeAppel = tym_Tous
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULBAP.TiersElipsis(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  wh  : string;
  Aux : string;
begin
  wh  := '(T_NATUREAUXI = "FOU" OR T_NATUREAUXI = "AUC" OR T_NATUREAUXI = "DIV")';
  Aux := THEdit(Sender).Text;
  if GetParamSocSecur('SO_CPMULTIERS', False) then begin
    Aux := AGLLanceFiche('CP','MULTIERS','','','M;' + Aux + ';' + Wh + ';' );
    if Aux <> '' then
      THEdit(Sender).Text := Aux;
  end
  else begin
    if Aux <> '' then
      wh := wh + ' AND (T_LIBELLE LIKE "' + Aux + '%" OR T_AUXILIAIRE LIKE "' + Aux + '%")';
    LookupList(THEdit(Sender), TraduireMemoire('Auxiliaire'), 'TIERS', 'T_AUXILIAIRE', 'T_LIBELLE', Wh, 'T_AUXILIAIRE', True, 2)
  end;
end;

{JP 13/05/08 : FQ 22665 : récapitulatif du bon à payer
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULBAP.BDetailOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  {$IFDEF EAGLCLIENT}
  F : THGrid;
  {$ELSE}
  F : THDBGrid;
  {$ENDIF}
  n : Integer;
  TobEtat : TOB;

    {------------------------------------------------------------------------}
    procedure _ChargeEnreg;
    {------------------------------------------------------------------------}
    var
      T : TOB;
    begin
      T := TOB.Create('MaTob', TobEtat, -1);
      //BAP_CODEVISA;BAP_CIRCUITBAP;BAP_CREATEUR;BAP_NUMEROORDRE;BAP_VISEUR1;BAP_STATUTBAP;BAP_JOURNAL;BAP_NUMEROPIECE;BAP_DATECOMPTABLE;BAP_NBJOUR;BAP_EXERCICE;BAP_BLOCNOTE;BAP_DATEECHEANCE;BAP_VISEUR;E_AUXILIAIRE;E_CREDIT;E_DEBIT;BAP_ECHEANCEBAP;
      T.AddChampSupValeur('BAP_JOURNAL', TFMul(Ecran).Q.FindField('BAP_JOURNAL').AsString);
      T.AddChampSupValeur('BAP_EXERCICE', TFMul(Ecran).Q.FindField('BAP_EXERCICE').AsString);
      T.AddChampSupValeur('BAP_DATECOMPTABLE', TFMul(Ecran).Q.FindField('BAP_DATECOMPTABLE').AsDateTime);
      T.AddChampSupValeur('BAP_NUMEROPIECE', TFMul(Ecran).Q.FindField('BAP_NUMEROPIECE').AsInteger);
      T.AddChampSupValeur('BAP_NUMEROORDRE', TFMul(Ecran).Q.FindField('BAP_NUMEROORDRE').AsInteger);
      T.AddChampSupValeur('BAP_CODEVISA', TFMul(Ecran).Q.FindField('BAP_CODEVISA').AsString);
      T.AddChampSupValeur('BAP_CIRCUITBAP', TFMul(Ecran).Q.FindField('BAP_CIRCUITBAP').AsString);
      T.AddChampSupValeur('BAP_CREATEUR', TFMul(Ecran).Q.FindField('BAP_CREATEUR').AsString);
      T.AddChampSupValeur('BAP_VISEUR1', TFMul(Ecran).Q.FindField('BAP_VISEUR1').AsString);
      T.AddChampSupValeur('BAP_STATUTBAP', TFMul(Ecran).Q.FindField('BAP_STATUTBAP').AsString);
      T.AddChampSupValeur('BAP_NBJOUR', TFMul(Ecran).Q.FindField('BAP_NBJOUR').AsInteger);
      T.AddChampSupValeur('BAP_BLOCNOTE', TFMul(Ecran).Q.FindField('BAP_BLOCNOTE').AsString);
      T.AddChampSupValeur('BAP_DATEECHEANCE', TFMul(Ecran).Q.FindField('BAP_DATEECHEANCE').AsDateTime);
      T.AddChampSupValeur('BAP_VISEUR', TFMul(Ecran).Q.FindField('BAP_VISEUR').AsString);
      T.AddChampSupValeur('E_AUXILIAIRE', TFMul(Ecran).Q.FindField('E_AUXILIAIRE').AsString);
      T.AddChampSupValeur('E_CREDIT', TFMul(Ecran).Q.FindField('E_CREDIT').AsFloat);
      T.AddChampSupValeur('E_DEBIT', TFMul(Ecran).Q.FindField('E_DEBIT').AsFloat);
      T.AddChampSupValeur('BAP_ECHEANCEBAP', TFMul(Ecran).Q.FindField('BAP_ECHEANCEBAP').AsDateTime);
    end;

begin
  {$IFDEF EAGLCLIENT}
  F := THGrid(TFMul(Ecran).FListe);
  {$ELSE}
  F := THDBGrid(TFMul(Ecran).FListe);
  {$ENDIF}

  {Aucune sélection, on sort}
  if (F.NbSelected = 0) and not F.AllSelected then begin
    HShowMessage('0;' + Ecran.Caption + ';Aucun élément n''est sélectionné.;W;O;O;O;', '', '');
    Exit;
  end;

  TobEtat := TOB.Create('_MaTob', nil, -1);
  try
    {$IFNDEF EAGLCLIENT}
    TFMul(Ecran).Q.First;
    if F.AllSelected then
      while not TFMul(Ecran).Q.EOF do begin
        _ChargeEnreg;
        TFMul(Ecran).Q.Next;
      end
    else
    {$ENDIF}

    for n := 0 to F.nbSelected - 1 do begin
      F.GotoLeBookmark(n);
      {$IFDEF EAGLCLIENT}
      TFMul(Ecran).Q.TQ.Seek(F.Row - 1);
      {$ENDIF}
      _ChargeEnreg;
    end;
    if TobEtat.Detail.Count > 0 then
      LanceEtatTob('E', 'BAP', 'REC', TobEtat, True, False, False, nil, '', TraduireMemoire('Détail des BAP'),False);
  finally
    FreeAndNil(TobEtat);
  end;
end;

initialization
  RegisterClasses([TOF_CPMULBAP]);

end.



