{ Unité : Source TOF du Mul : CPMULCIRCUITBAP ()
--------------------------------------------------------------------------------------
    Version    |   Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 7.01.001.001   02/02/06   JP   Création de l'unité
 8.01.001.001   22/01/07   JP   FQ 19442 : filtrer les utilisateurs sur les tablettes
 8.00.001.018   29/05/07   JP   Mise en place des rôles à la place des groupes pour filtrer les utilisateurs
 8.10.002.001   18/10/07   JP   FQ 19875 : mise en place d'un utilitaire de contrôle des utilisateurs
--------------------------------------------------------------------------------------}
unit CPMULCIRCUITBAP_TOF;

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
  SysUtils, UTob, uTOF, HCtrls;

  type
  TOF_CPMULCIRCUITBAP = class (TOF)
    procedure OnArgument(S : string); override;
  private
    {$IFDEF EAGLCLIENT}
    FListe : THGrid;
    {$ELSE}
    FListe : THDBGrid;
    {$ENDIF EAGLCLIENT}

    procedure BInsertClick(Sender : TObject);
    procedure BOuvrirClick(Sender : TObject);
    procedure BDeleteClick(Sender : TObject);
    procedure SlctAllClick(Sender : TObject);
    procedure ElipsisClick(Sender : TObject);
    procedure BRepareClick(Sender : TObject); {18/10/07 : FQ 19875}
  end;

procedure CpLanceFiche_CircuitMul;

implementation

uses
  HEnt1, HMsgBox, Ent1, HTB97, CPCIRCUITBAP_TOF, uLibBonAPayer, ParamSoc, CPCONTROLEBAP_TOF;

{---------------------------------------------------------------------------------------}
procedure CpLanceFiche_CircuitMul;
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche('CP', 'CPMULCIRCUITBAP', '', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCIRCUITBAP.OnArgument(S: string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 7509020;
  FListe := TFMul(Ecran).FListe;
  TToolbarButton97(GetControl('BINSERT'   )).OnClick := BInsertClick;
  TToolbarButton97(GetControl('BOUVRIR'   )).OnClick := BOuvrirClick;
  TToolbarButton97(GetControl('BDELETE'   )).OnClick := BDeleteClick;
  TToolbarButton97(GetControl('BSELECTALL')).OnClick := SlctAllClick;
  TToolbarButton97(GetControl('BREPARE'   )).OnClick := BRepareClick; {18/10/07 : FQ 19875}
  FListe.OnDblClick := BOuvrirClick;
  {JP 22/01/07 : FQ 19442 : on filtre les utilisateurs sur les seuls viseurs :
                 Usergroup qui correspond au paramsoc dans l'onglet tiers
                 adresse @mail renseignée dans la fiche utilsateur
   JP 29/05/07 : Mise en place des rôles}
  if GetControl('CCI_VISEUR1') is THEdit then begin
    (GetControl('CCI_VISEUR1') as THEdit).OnElipsisClick := ElipsisClick;
    (GetControl('CCI_VISEUR2') as THEdit).OnElipsisClick := ElipsisClick;
  end
  else begin
    (GetControl('CCI_VISEUR1') as THValComboBox).Plus := 'US_GROUPE = "' + GetParamSocSecur('SO_CPGROUPEVISEUR', '') + '" AND (US_EMAIL <> "" AND US_EMAIL IS NOT NULL)';
    (GetControl('CCI_VISEUR2') as THValComboBox).Plus := 'US_GROUPE = "' + GetParamSocSecur('SO_CPGROUPEVISEUR', '') + '" AND (US_EMAIL <> "" AND US_EMAIL IS NOT NULL)';
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCIRCUITBAP.BInsertClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  CpLanceFiche_CircuitFiche('', 'C');
  {Raffraîchissement de la liste}
  TToolbarButton97(GetControl('BCHERCHE')).Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCIRCUITBAP.BOuvrirClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  CpLanceFiche_CircuitFiche(VarToStr(GetField('CCI_CIRCUITBAP')), 'M');
  {Raffraîchissement de la liste}
  TToolbarButton97(GetControl('BCHERCHE')).Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCIRCUITBAP.SlctAllClick(Sender : TObject);
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
procedure TOF_CPMULCIRCUITBAP.BDeleteClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  p : Integer;
  {$IFDEF EAGLCLIENT}
  F : THGrid;
  {$ELSE}
  F : THDBGrid;
  {$ENDIF}

    {------------------------------------------------------------------------}
    procedure DetruireCircuit;
    {------------------------------------------------------------------------}
    var
      Code : string;
    begin
      Code := TFMul(Ecran).Q.FindField('CCI_CIRCUITBAP').AsString;
      if CanDeleteCircuit(Code) then begin
        ExecuteSQL('DELETE FROM CPCIRCUIT WHERE CCI_CIRCUITBAP = "' + Code + '"');
        Inc(p);
      end
      else
        HShowMessage('0;' + Ecran.Caption + ';Le type visa "' + Code  + '" est utilisé dans les bons à payer'#13 +
                     'et ne peut être supprimé.;I;O;O;O;', '', '');
    end;

begin
  {$IFDEF EAGLCLIENT}
  F := THGrid(TFMul(Ecran).FListe);
  {$ELSE}
  F := THDBGrid(TFMul(Ecran).FListe);
  {$ENDIF}

  p := 0;

  {Aucune sélection, on sort}
  if (F.NbSelected = 0) and not F.AllSelected then begin
    HShowMessage('0;' + Ecran.Caption + ';Aucun élément n''est sélectionné.;W;O;O;O;', '', '');
    Exit;
  end
  else if HShowMessage('0;' + Ecran.Caption + ';Êtes vous sûr de vouloir supprimer la sélection ?;Q;YN;N;N;', '', '') <> mrYes then
    Exit;

  BeginTrans;
  try
    {$IFNDEF EAGLCLIENT}
    TFMul(Ecran).Q.First;
    if F.AllSelected then
      while not TFMul(Ecran).Q.EOF do begin
        DetruireCircuit;
        TFMul(Ecran).Q.Next;
      end
    else
    {$ENDIF}

    for n := 0 to F.nbSelected - 1 do begin
      F.GotoLeBookmark(n);
      {$IFDEF EAGLCLIENT}
      TFMul(Ecran).Q.TQ.Seek(F.Row - 1);
      {$ENDIF}
      DetruireCircuit;
    end;
    CommitTrans;
  except
    on E : Exception do begin
      RollBack;
      HShowMessage('0;' + Ecran.Caption + '; Traitement interrompu :'#13 + E.Message + ';E;O;O;O;', '', '');
    end;
  end;

   HShowMessage('0;' + Ecran.Caption + ';' + IntToStr(p) + ' circuit(s) a (ont) été supprimé(s).;I;O;O;O;', '', '');
  {Raffraîchissement de la liste}
  TToolbarButton97(GetControl('BCHERCHE')).Click;
end;

{JP 29/05/07 : Mise en place des rôles
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCIRCUITBAP.ElipsisClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  LookUpUtilisateur(Sender);
end;

{JP 18/10/07 : FQ 19875 : mise en place d'un utilitaire de contrôle des utilisateurs ayant
               quitté la société (n'existant plus dans la table UTILISAT)
{---------------------------------------------------------------------------------------}
procedure TOF_CPMULCIRCUITBAP.BRepareClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  F : TOB;
  S : string;
  B : string;
begin
  {Requête de récupération des circuits ayant des utilisateurs "absents"}
  S := 'SELECT CCI_CIRCUITBAP, CCI_LIBELLE , CCI_NUMEROORDRE, CCI_VISEUR1, CCI_VISEUR2, ' +
       'U1.US_UTILISATEUR AS CCI_LIBVISEUR1, U2.US_UTILISATEUR AS CCI_LIBVISEUR2 FROM CPCIRCUIT ';
  S := S + 'LEFT JOIN UTILISAT U1 ON U1.US_UTILISATEUR = CCI_VISEUR1 ' +
           'LEFT JOIN UTILISAT U2 ON U2.US_UTILISATEUR = CCI_VISEUR2 ' +
           'WHERE U1.US_UTILISATEUR IS NULL OR U2.US_UTILISATEUR IS NULL';

  {Requête de récupération des bons à payer ayant des utilisateurs "absents"}
  B := 'SELECT BAP_JOURNAL, BAP_DATECOMPTABLE, BAP_NUMEROPIECE, BAP_ECHEANCEBAP, BAP_CIRCUITBAP, BAP_VISEUR1, BAP_VISEUR2, ' +
       'BAP_EXERCICE, U1.US_LIBELLE AS BAP_LIBVISEUR1, U2.US_LIBELLE AS BAP_LIBVISEUR2, "-" AS BAP_MODIF, BAP_NUMEROORDRE FROM CPBONSAPAYER ';
  B := B + 'LEFT JOIN UTILISAT U1 ON U1.US_UTILISATEUR = BAP_VISEUR1 ' +
           'LEFT JOIN UTILISAT U2 ON U2.US_UTILISATEUR = BAP_VISEUR2 ' +
           'WHERE (U1.US_UTILISATEUR IS NULL OR U2.US_UTILISATEUR IS NULL) ';
  B := B + 'AND NOT (BAP_STATUTBAP IN ("' + sbap_Valide + '","' + sbap_Definitif + '"))';

  T := TOB.Create('mmmm', nil, -1);
  F := TOB.Create('bbbb', nil, -1);
  try
    T.LoadDetailFromSQL(S);
    F.LoadDetailFromSQL(B);
    if (F.Detail.Count = 0) and (T.Detail.Count = 0) then
      PgiInfo(TraduireMemoire('Les circuits et les bons à payer sont cohérents avec la table "UTILISAT".'), Ecran.Caption)
    else
      CpLanceFiche_ControlBap(T, F);

  finally
    FreeAndNil(T);
    FreeAndNil(F);
  end;
end;

initialization
  RegisterClasses([TOF_CPMULCIRCUITBAP]);

end.
