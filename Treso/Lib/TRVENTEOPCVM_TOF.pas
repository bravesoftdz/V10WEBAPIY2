{-------------------------------------------------------------------------------------
  Version    |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 6.2X.001.001  04/11/04  JP  Création de l'unité : vente multiple / simple d'OPCVM
 8.01.001.001  15/12/06  JP  FQ 10373 : On passe le nombre de parts d'un Integer à un Double
 8.01.001.014  09/05/07  JP  FQ 10451 : Modification du contrôle de la date de vente / à l'achat
--------------------------------------------------------------------------------------}
unit TRVENTEOPCVM_TOF;


interface

uses {$IFDEF VER150} variants,{$ENDIF}
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_MAIN,
  {$ENDIF}
  StdCtrls, Controls, Classes, Forms, SysUtils, HCtrls, HEnt1, UTOF,
  UObjOPCVM, uTob;

type
  TOF_TRVENTEOPCVM = class(TOF)
    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
    procedure OnUpdate              ; override;
  private
    ObjVente : TObjOPCVMVente;
    {Routine de lancement du calcul : le traitement est radicalement différent
     entre une vente simple et une vente multiple}
    procedure CalculerVente(T : TOB; VideTob : Boolean);
  protected
    lOpcvm   : TOB;
    Grille   : THGrid;
    EntreOk  : Boolean;
    AnnuleOk : Boolean; {09/05/07 / FQ 10451}
    {évènements sur les boutons}
    procedure BImprimerClick  (Sender : TObject);
    procedure BSupprimeClick  (Sender : TObject);
    procedure BCalculerClick  (Sender : TObject);
    {évènements sur la grille}
    procedure GrilleBtnClick  (Sender : TObject);
    procedure GrilleDblClick  (Sender : TObject);
    procedure GrilleEnter     (Sender : TObject);
    procedure GrilleCellExit  (Sender : TObject; var aCol, aRow : Integer; var Cancel : Boolean);
    procedure GrilleCellEnter (Sender : TObject; var aCol, aRow : Integer; var Cancel : Boolean);
    procedure GrilleCellSelect(Sender : TObject; ACol, ARow: Longint; var CanSelect: Boolean);
    procedure GrilleKeyDown   (Sender : TObject; var Key : Word; Shift : TShiftState);
    procedure GrilleKeyPress  (Sender: TObject; var Key: Char);
  public
    Multiple  : string;

    procedure GererAffichage   ;
    procedure AssigneProprietes;
    procedure TaillerGrille    ;
    procedure OnAfterFormShow  ;
    procedure SupprimerVente(var T : TOB; Detruire : Boolean);
  end;

procedure TRLanceFiche_VenteOPCVM(Dom , Fiche, Range, Lequel, Arguments : string);


implementation

uses
  {$IFNDEF EAGLCLIENT}
  PrintDBG,
  {$ENDIF}
  HMsgBox, Vierge, HTB97, Constantes, Grids, Commun, Messages,
  ExtCtrls, Windows, ParamDat, AglInit, TROPCVM_TOM, uTobDebug;

const
  COL_NUMOPCVM     = 0;
  COL_CODEOPCVM    = 1;
  COL_TRANSACTION  = 2;
  COL_GENERAL      = 3;
  COL_LIBELLE      = 4;
  COL_BASECALCUL   = 5;
  COL_MONTANTACH   = 6;
  COL_FRAISACH     = 7;
  COL_TVAACHAT     = 8;
  COL_COMACHAT     = 9;
  COL_COTATIONACH  = 10;
  COL_NBPARTACHETE = 11;
  COL_NBPARTVENDU  = 12;
  COL_MONTANTTOT   = 13;
  COL_DATEVENTE    = 14;
  COL_PARTAVENDRE  = 15;
  COL_MONTANTVEN   = 16;
  COL_DATEACHAT    = 17;


{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_VenteOPCVM(Dom , Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  ch : string;
begin
  inherited;
  Ecran.HelpContext := 150;
  ch := ReadTokenSt(S);
  Multiple := ch;
  {Affection des propriétés et des évènements des composant de la fiche}
  AssigneProprietes;
  {Gestion de l'ergonomie}
  GererAffichage;
  if (Multiple = vop_FIFO) or (Multiple = vop_CAMP) then
    Grille.ColFormats[COL_NUMOPCVM] := 'CB=TRPORTEFEUILLE||';
  {Gestion de la Grille}
  Grille.ColCount := COL_DATEACHAT + 1;
  {18/12/06 : FQ 10373 : Le nombre de part peut être saisi au millième. Pour que le format s'applique
              il faut qu'il soit mis avant le PutGridDetail}
  Grille.ColFormats[COL_PARTAVENDRE ] := '#,##0.000';
  Grille.ColFormats[COL_NBPARTACHETE] := '#,##0.000';
  Grille.ColFormats[COL_NBPARTVENDU ] := '#,##0.000';
  lOpcvm.PutGridDetail(Grille, True, False, '');
  TaillerGrille;
  {Création de l'objet qui va calculer les lignes de ventes}
  ObjVente := TObjOPCVMVente.Create(Multiple);
  AnnuleOk := False;

  TFVierge(Ecran).OnAfterFormShow := OnAfterFormShow; 
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(ObjVente) then FreeAndNil(ObjVente);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.OnUpdate;
{---------------------------------------------------------------------------------------}
var
  c, r   : Integer;
  Cancel :  Boolean;
begin
  inherited;
  c := Grille.Col;
  r := Grille.Row;
  Cancel := False;
  {Le SpeedButton, ne prenant pas le focus, on force le traitement}
  GrilleCellExit(Grille, c, r, Cancel);

  if ObjVente.EnAttente then begin
    BeginTrans;
    try
      ObjVente.PutVenteInBase;
      ObjVente.tOPCVM.UpdateDB(True);
      CommitTrans;
      {On ne le repasse à False que si tout s'est bien passé !!!}
      ObjVente.EnAttente := False;
    except
      on E : Exception do begin
        RollBack;
        LastError := 2;
        LastErrorMsg := E.Message;
      end;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.GererAffichage;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF EAGLCLIENT}
  SetControlVisible('BIMPRIMER', False);
  {$ENDIF}
  SetControlVisible('BINSERT', False);

  {Si on est en vente multiple}
  if Multiple =  vop_FIFO then begin
    Ecran.Caption := 'Vente multiple (FIFO)';
    UpdateCaption(Ecran);
  end
  else if Multiple =  vop_CAMP then begin
    Ecran.Caption := 'Vente multiple (CAMP)';
    //SetControlVisible('BCALCUL', True);
    UpdateCaption(Ecran);
  end
  else if Multiple = vop_Autre then begin
    Ecran.Caption := 'Détail d''une vente multiple';
    UpdateCaption(Ecran);
    SetControlVisible('BDELETE' , False);
    SetControlVisible('BVALIDER', False);
  end
  else
    SetControlVisible('BCALCUL', True);

  {Si par extraordinaire la monnaie de tenue n'est pas l'Euro}
  if V_PGI.DevisePivot <> 'EUR' then begin
    AssignDrapeau(Timage(GetControl('IDEV')), V_PGI.DevisePivot);
    SetControlCaption('LBDEVISE', 'Les montants sont exprimés en ' + RechDom('TTDEVISE', V_PGI.DevisePivot, False));
    Timage(GetControl('IDEV')).Left := THLabel(GetControl('LBDEVISE')).Left + THLabel(GetControl('LBDEVISE')).Width + 5;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.AssigneProprietes;
{---------------------------------------------------------------------------------------}
begin
  lOpcvm := TFVierge(Ecran).LaTOF.LaTOB;
  Grille := THGrid(GetControl('GRID'));
  TToolbarButton97(GetControl('BIMPRIMER')).OnClick := BImprimerClick;
  TToolbarButton97(GetControl('BDELETE'  )).OnClick := BSupprimeClick;
  TToolbarButton97(GetControl('BCALCUL'  )).OnClick := BCalculerClick;
  Grille.OnCellEnter     := GrilleCellEnter;
  Grille.OnCellExit      := GrilleCellExit;
  Grille.OnSelectCell    := GrilleCellSelect;
  Grille.OnElipsisClick  := GrilleBtnClick;
  Grille.OnKeyDown       := GrilleKeyDown;
  Grille.OnKeyPress      := GrilleKeyPress;
  Grille.OnDblClick      := GrilleDblClick;
  Grille.OnEnter         := GrilleEnter;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.SupprimerVente(var T : TOB; Detruire : Boolean);
{---------------------------------------------------------------------------------------}
var
  F : TOB;
begin
  {On recherche une éventuelle ligne de vente (TRVENTEOPCVM) pour la supprimer}
  if (Multiple = vop_FIFO) or (Multiple = vop_CAMP) then begin
    if Assigned(ObjVente.tVente) then begin
      {En vente mutliple, X achats sont rassemblés pour une ligne de la grille}
      F := ObjVente.tVente.FindFirst(['TVE_CODEOPCVM', 'TVE_PORTEFEUILLE', 'TVE_GENERAL'],
          [T.GetString('TOP_CODEOPCVM'), T.GetString('TOP_PORTEFEUILLE'), T.GetString('TOP_GENERAL')], True);
      while F <> nil do begin
        {Si on a trouvé une vente, on la supprime}
        FreeAndNil(F);
        F := ObjVente.tVente.FindNext(['TVE_CODEOPCVM', 'TVE_PORTEFEUILLE', 'TVE_GENERAL'],
            [T.GetString('TOP_CODEOPCVM'), T.GetString('TOP_PORTEFEUILLE'), T.GetString('TOP_GENERAL')], True);
      end;
    end;

    if Assigned(ObjVente.tOPCVM) then begin
      {Suppression des tobs de mise à jour de la table TROPCVM}
      F := ObjVente.tOPCVM.FindFirst(['TOP_CODEOPCVM', 'TOP_PORTEFEUILLE', 'TOP_GENERAL'],
          [T.GetString('TOP_CODEOPCVM'), T.GetString('TOP_PORTEFEUILLE'), T.GetString('TOP_GENERAL')], True);
      while F <> nil do begin
        {Si on a trouvé une vente, on la supprime}
        FreeAndNil(F);
        F := ObjVente.tOPCVM.FindNext(['TOP_CODEOPCVM', 'TOP_PORTEFEUILLE', 'TOP_GENERAL'],
            [T.GetString('TOP_CODEOPCVM'), T.GetString('TOP_PORTEFEUILLE'), T.GetString('TOP_GENERAL')], True);
      end;
    end;
  end

  {Vente simple}
  else begin
    {En vente simple, un achat = une vente, on travaille sur TOP_NUMOPCVM}
    F := ObjVente.tVente.FindFirst(['TVE_NUMTRANSAC'], [T.GetString('TOP_NUMOPCVM')], True);
    {Si on a trouvé une vente, on la supprime}
    if Assigned(F) then FreeAndNil(F);
    {Suppression des tobs de mise à jour de la table TROPCVM}
    F := ObjVente.tVente.FindFirst(['TOP_NUMTRANSAC'], [T.GetString('TOP_NUMOPCVM')], True);
    {Si on a trouvé une Tob de Maj, on la supprime}
    if Assigned(F) then FreeAndNil(F);
  end;

  {On détruit la tob OPCVM courante}
  if Detruire then FreeAndNil(T);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.TaillerGrille;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Grille.TwoColors := True;
  {Largeurs des colonnes}
  Grille.ColWidths[COL_NUMOPCVM    ] := 100;
  Grille.ColWidths[COL_CODEOPCVM   ] := 58;
  Grille.ColWidths[COL_TRANSACTION ] := -1;
  Grille.ColWidths[COL_GENERAL     ] := 58;
  Grille.ColWidths[COL_LIBELLE     ] := 110;
  Grille.ColWidths[COL_BASECALCUL  ] := -1;
  Grille.ColWidths[COL_MONTANTACH  ] := -1;
  Grille.ColWidths[COL_FRAISACH    ] := -1;
  Grille.ColWidths[COL_TVAACHAT    ] := -1;
  Grille.ColWidths[COL_COMACHAT    ] := -1;
  Grille.ColWidths[COL_COTATIONACH ] := -1;
  Grille.ColWidths[COL_NBPARTVENDU ] := 45;
  Grille.ColWidths[COL_NBPARTACHETE] := 45;
  Grille.ColWidths[COL_MONTANTTOT  ] := 75;
  Grille.ColWidths[COL_PARTAVENDRE ] := 45;
  Grille.ColWidths[COL_DATEVENTE   ] := 74;
  Grille.ColWidths[COL_MONTANTVEN  ] := 75;
  Grille.ColWidths[COL_DATEACHAT   ] := -1;
  {Titre des colonnes}
  if (Multiple = vop_FIFO) or (Multiple = vop_CAMP) then
    Grille.Cells[COL_NUMOPCVM  , 0] := 'Portefeuille'
  else
    Grille.Cells[COL_NUMOPCVM  , 0] := 'Numéro';
  Grille.Cells[COL_CODEOPCVM   , 0] := 'Code';
  Grille.Cells[COL_TRANSACTION , 0] := '';
  Grille.Cells[COL_GENERAL     , 0] := 'Compte';
  Grille.Cells[COL_LIBELLE     , 0] := 'Libellé';
  Grille.Cells[COL_BASECALCUL  , 0] := '';
  Grille.Cells[COL_MONTANTACH  , 0] := '';
  Grille.Cells[COL_FRAISACH    , 0] := '';
  Grille.Cells[COL_TVAACHAT    , 0] := '';
  Grille.Cells[COL_COMACHAT    , 0] := '';
  Grille.Cells[COL_COTATIONACH , 0] := '';
  Grille.Cells[COL_NBPARTACHETE, 0] := 'Achetés';
  Grille.Cells[COL_NBPARTVENDU , 0] := 'Vendus';
  Grille.Cells[COL_MONTANTTOT  , 0] := 'Mnt Achat';
  Grille.Cells[COL_PARTAVENDRE , 0] := 'A vend.';
  Grille.Cells[COL_DATEVENTE   , 0] := 'Date vente';
  Grille.Cells[COL_MONTANTVEN  , 0] := 'Mnt Vente';
  Grille.Cells[COL_DATEACHAT   , 0] := '';
  {Alignement dans les cellules}
  Grille.ColAligns[COL_NBPARTVENDU ] := taRightJustify;
  Grille.ColAligns[COL_NBPARTACHETE] := taRightJustify;
  Grille.ColAligns[COL_MONTANTTOT  ] := taRightJustify;
  Grille.ColAligns[COL_PARTAVENDRE ] := taRightJustify;
  Grille.ColAligns[COL_MONTANTVEN  ] := taRightJustify;
  Grille.ColAligns[COL_CODEOPCVM   ] := taCenter;
  Grille.ColAligns[COL_GENERAL     ] := taCenter;
  if (Multiple <> vop_FIFO) and (Multiple <> vop_CAMP) then
    Grille.ColAligns[COL_NUMOPCVM  ] := taCenter;

  {Saisie dans la grille}
  Grille.ColTypes  [COL_DATEVENTE] := 'D';
  Grille.ColFormats[COL_DATEVENTE] := ShortDateFormat;

  if Multiple = vop_Autre then
    Grille.Options := Grille.Options + [goRowSelect];

  for n := 0 to Grille.ColCount - 1 do begin
    if (Multiple = vop_FIFO) or (Multiple = vop_CAMP) then
      Grille.ColEditables[n] := n in [COL_PARTAVENDRE, COL_DATEVENTE]
    else if Multiple = vop_Simple then
      Grille.ColEditables[n] := n = COL_DATEVENTE;

    if n in [COL_TRANSACTION, COL_BASECALCUL, COL_MONTANTACH, COL_FRAISACH, COL_TVAACHAT,
             COL_COMACHAT, COL_COTATIONACH] then Grille.ColLengths[n] := -1;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.GrilleCellEnter(Sender : TObject; var aCol, aRow : Integer; var Cancel : Boolean);
{---------------------------------------------------------------------------------------}
begin
  {Pour ne pas voir la combo, car la colonne n'est pas éditable}
  if ((Multiple = vop_FIFO) or (Multiple = vop_CAMP)) and (Grille.Col = COL_NUMOPCVM) then
    Grille.ValCombo.Visible := False;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.GrilleCellExit(Sender : TObject; var aCol, aRow : Integer; var Cancel : Boolean);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  N : Double;
begin
  if Multiple = vop_Autre then Exit;

  {Traitement de la date de vente}
  if aCol = COL_DATEVENTE then begin
    T := TOB(Grille.Objects[0, ARow]);

    if Assigned(T) then begin
      {On s'assure que la date est correcte}
      if IsValidDate(Grille.CellValues[aCol, aRow]) and
         (VarToDateTime(Grille.CellValues[aCol, aRow]) >= 1) then begin
        if (VarToDateTime(Grille.CellValues[aCol, aRow]) >= VarToDateTime(Grille.CellValues[COL_DATEACHAT, aRow])) then
          {Mise à jour de la Tob}
          T.SetDateTime('DATEVENTE', StrToDateTime(Grille.Cells[aCol, aRow]))
        else begin
          Cancel := True;
          PGIError(TraduireMemoire('La date ne peut être inférieure à celle d''achat.'), Ecran.Caption);
        end;
      end
      else begin
        {Sinon on interdit la sortie de la cellule}
        Cancel := True;
        PGIError(TraduireMemoire('La date saisie n''est pas correcte.'), Ecran.Caption);
      end;
    end {Assigned(T)}

    {Si on n'arrive pas à récupérer la Tob}
    else
      PGIError(TraduireMemoire('Une erreur est intervenue : la vente ne pourra pas être créée.'), Ecran.Caption);
  end

  {Traitement du nombre de parts à vendre}
  else if (Multiple <> vop_Simple) and (aCol = COL_PARTAVENDRE) then begin
    T := TOB(Grille.Objects[0, ARow]);
    {On récupère la valeur saisie
    15/12/06 : FQ 10373 : On passe d'un Integer à un Double}
    N := Valeur(Grille.CellValues[aCol, aRow]);

    if Assigned(T) then begin
      {On interdit la saisie d'un nombre plus grand que la différence entre le nombre
       de parts achetées et le nombre de parts vendues}
      if N > (T.GetDouble('TOP_NBPARTACHETE') - T.GetDouble('TOP_NBPARTVENDU')) then begin
        Cancel := True;
        PGIError(TraduireMemoire('Le nombre de parts saisies est trop important.'), Ecran.Caption);
      end
      else begin
        {Mise à jour de la tob}
        T.SetDouble('PARTAVENDRE', N);

        {Calcule de la vente FIFO}
        if (Multiple = vop_FIFO) or (Multiple = vop_CAMP) then begin
          {09/05/07 : FQ 10451 : Pour éviter un nouveau calcul en cas d'annulation de la vente}
          if not AnnuleOk then begin
            CalculerVente(T, False);
            {Mise à jour de la gille}
            Grille.Cells[COL_MONTANTVEN, aRow] := T.GetString('MONTANTVEN');
            {09/05/07 : FQ 10451 : on rend le focus à la date si le montant n'a pas pu être calculé}
            if T.GetDouble('MONTANTVEN') = 0 then begin
              {Pn annule le nombre de parts à vendre}
              Grille.Cells[COL_PARTAVENDRE, aRow] := T.GetString('PARTAVENDRE');
              Grille.Col := COL_DATEVENTE;
            end;
          end
          else
            AnnuleOk := False;
        end;
      end;
    end{Assigned(T)}

    {Si on n'arrive pas à récupérer la Tob}
    else
      PGIError(TraduireMemoire('Une erreur est intervenue : la vente ne pourra pas être créée.'), Ecran.Caption);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.GrilleKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  if Multiple = vop_Autre then Exit;

  if (Grille.Col = COL_DATEVENTE) then begin
    if (Key = VK_F5) then
      {Appel de la fiche de saisie de date}
      V_PGI.ParamDateProc(Grille)
  end
  else if Key = VK_SPACE then
    {Je suis obligé d'enlevé goRowSelect pour gérér le goEditing dans les options de la
     grille : donc je gère la gestion de la sélection à la main}
    Grille.FlipSelection(Grille.Row);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.GrilleKeyPress(Sender: TObject; var Key: Char);
{---------------------------------------------------------------------------------------}
begin
  if Multiple = vop_Autre then Exit;
  if (Grille.Col = COL_DATEVENTE) then
    {Gestion des raccourcis agl sur les dates}
    ParamDate(Ecran, Grille, Key);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.GrilleCellSelect(Sender: TObject; ACol, ARow: Longint; var CanSelect: Boolean);
{---------------------------------------------------------------------------------------}
begin
  {Affichage de l'elipsis si on est sur la colonne COL_DATEVENTE}
  Grille.ElipsisButton := ACol = COL_DATEVENTE;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.GrilleBtnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Appel de la fiche de saisie de date}
  V_PGI.ParamDateProc(Grille);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.GrilleEnter(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if not EntreOk then begin
    {Pour ne pas relancer le code}
    EntreOk := True;
    {Pour exécuter le CellEnter en entrée de la Grille et ainsi cacher ValCombo}
    PostMessage(Grille.Handle, WM_KEYDOWN, VK_TAB, 0);
    PostMessage(Grille.Handle, WM_KEYDOWN, VK_LEFT, 0);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.GrilleDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  F : TOB;
  n : Integer;
  D : Double;
begin
  {Dans le cas d'une vente multiple, pour chaque ligne de la grille peut correspondre plusieurs
   enregistrements dans la table}
  if (Multiple = vop_FIFO) or (Multiple = vop_CAMP) then begin
    T := TOB.Create('ùùù', nil, -1);
    try
      {On va chercher dans la table TROPCVM les transactions correspondant à cette clef}
      ChargeDetailOPCVM(T, Grille.CellValues[COL_NUMOPCVM , Grille.Row],
                           Grille.CellValues[COL_CODEOPCVM, Grille.Row],
                           Grille.CellValues[COL_GENERAL  , Grille.Row]);

      {Si des ventes on été calculées, on met à jour les champs PARTAVENDRE et MONTANTVEN.
       Remarque : on ne peut assigner ObjVente.tOPCVM à T, car ObjVente.tOPCVM ne contient
                  que les enregistrements sur lesquels il y a une vente. il y a de forts
                  risques que tous les enregistements de la clef ne soient pas concernés !}
      if ObjVente.tOPCVM.Detail.Count > 0 then begin
        for n := 0 to T.Detail.Count - 1 do begin
          F := ObjVente.tOPCVM.FindFirst(['TOP_NUMOPCVM'], [T.Detail[n].GetString('TOP_NUMOPCVM')], True);
          if Assigned(F) then begin
            {15/12/06 : FQ 10373 : On passe d'un Integer à un Double}
            T.Detail[n].SetDouble('PARTAVENDRE', F.GetDouble('NBPARTS'));
            {Le montant de la vente est la différence entre les montants de F et T.Detail[n] qui contient
             les montants avant calcul}
            D := F.GetDouble('TOP_FRAISVEN')   - T.Detail[n].GetDouble('TOP_FRAISVEN') +
                 F.GetDouble('TOP_TVAVENTE')   - T.Detail[n].GetDouble('TOP_TVAVENTE') +
                 F.GetDouble('TOP_COMVENTE')   - T.Detail[n].GetDouble('TOP_COMVENTE') +
                 F.GetDouble('TOP_MONTANTVEN') - T.Detail[n].GetDouble('TOP_MONTANTVEN');
            T.Detail[n].SetDouble('MONTANTVEN', D);
          end;
        end;
      end;
      TheTob := T;
      TRLanceFiche_VenteOPCVM('TR', 'TRVENTEOPCVM', '', '', vop_Autre);
    finally
      FreeAndNil(T);
    end;
  end

  else
    {Il s'agit d'un affichage simple, on appelle la fiche}
    TRLanceFiche_OPCVM('TR', 'TROPCVM', '', Grille.Cells[COL_NUMOPCVM, Grille.Row], ActionToString(taConsult));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.BImprimerClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {$IFNDEF EAGLCLIENT}
  PrintDBGrid(Grille, nil, 'Liste des ventes', '');
  {$ENDIF}
end;

{Suppression de lignes du traitement
{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.BSupprimeClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  n : Integer;
begin
  if Grille.nbSelected = 0 then
    HShowMessage('0;' + Ecran.Caption + ';Aucune transaction n''est sélectionnée.;I;O;O;O;', '', '')
  else begin
    if HShowMessage('1;' + Ecran.Caption + ';Êtes-vous sûr de vouloir retirer les lignes sélectionnées.;Q;YNC;N;C;', '', '') = mrYes then begin
      for n := Grille.RowCount - 1 downto 1 do
        {Si la ligne est sélectionnée}
        if Grille.IsSelected(n) then begin

          T := TOB(Grille.Objects[0, n]);
          {Suppression de la tob et des éventulles ventes liées}
          if Assigned(T) then SupprimerVente(T, True);

          {On supprime la rangée}
          Grille.DeleteRow(Grille.Row);
        end;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.BCalculerClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  c, r   : Integer;
  Cancel :  Boolean;
begin
  c := Grille.Col;
  r := Grille.Row;
  Cancel := False;
  {Le SpeedButton, ne prenant pas le focus, on force le traitement}
  GrilleCellExit(Grille, c, r, Cancel);

  {On vide la Tob OPCVM de l'objet}
  ObjVente.tOPCVM.ClearDetail;
  for n := 0 to lOpcvm.Detail.Count - 1 do begin
    if n = 0 then CalculerVente(lOpcvm.Detail[n], True)
             else CalculerVente(lOpcvm.Detail[n], False);
  end;
  {Si on est en CAMP, on lance le calcul des plus / moins values et des rendements}
  if Multiple = vop_CAMP then
    ObjVente.PmvEtRendementCamp;

  {On vide la grille}
  Grille.VidePile(False);
  {On met à jour la grille}
  lOpcvm.PutGridDetail(Grille, True, False, '');
  {Dessin de la Grille}
  TaillerGrille;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.CalculerVente(T : TOB; VideTob : Boolean);
{---------------------------------------------------------------------------------------}
var
  F : TOB;
  O : TOB;
  n : Integer;
  p : Double;
  Obj : TObjCalculPMVR;
begin
  p := 0;
  {Vente simple : pour chaque ligne de la tob grille (lOpcvm), on charge l'enregistrement correspondant
                  dans la table TROPCVM, on génère l'enregistrement de vente à la date saisie dans la
                  grille, puis on met à jour la tob de la grille avec le montant calculé de la vente

   Vente CAMP   : Le principe est le même qu'une vente simple à l'exception du calcul des plus values
                  et des rendements qui sont calculés en fin de traitement afin de pouvoir calculer
                  le ratio CAMP}
  if (Multiple = vop_Simple) {or (Multiple = vop_CAMP)} then begin
    {Comme les transactions disparaissent de la liste qu'une fois fermées, on peut se retrouver
     avec des lignes où il n'y a rien à vendre, en plus des cas ou la saisie est à zéro : en tout
     état de cause, on ne traite pas ces lignes
     15/12/06 : FQ 10373 : On passe d'un Integer à un Double}
    if T.GetDouble('PARTAVENDRE') = 0 then Exit;

    F := TOB.Create('TROPCVM', ObjVente.tOPCVM, -1);
    F.SetString('TOP_NUMOPCVM', T.GetString('TOP_NUMOPCVM'));
    F.LoadDB;
    {09/05/07 : FQ 10451 : Correction du contrôle des dates de vente et d'achat. La vue TRVENTEMULTIOPCVMMS a été
                modifié pour ne prendre en compte la plus vieille date d'achat et non plus la plus récente = >
                le test dans le CellExit perd de son intérêts, quoiqu'il peut y avoir des gens distraits.
                Ici, en vente simple, le test ne semble pas non plus extrêmement utile}
    if T.GetDateTime('DATEVENTE') < F.GetDateTime('TOP_DATEACHAT') then begin
      PGIError(TraduireMemoire('La date ne peut être inférieure à celle d''achat.'), Ecran.Caption);
      T.SetDouble('MONTANTVEN', 0);
      T.SetDouble('PARTAVENDRE', 0);
      AnnuleOk := True;
      Exit;
    end;

    {15/12/06 : FQ 10373 : On passe d'un Integer à un Double}
    F.AddChampSupValeur('NBPARTS', T.GetDouble('PARTAVENDRE'));
    if ObjVente.GenererVente(F, VideTob, T.GetDateTime('DATEVENTE'), p) then
      {On récupère le montant total de la vente dans la dernière fille (ie celle qui vient d'être
       créée) de la Tob vente de l'objet}
      T.SetDouble('MONTANTVEN', ObjVente.tVente.Detail[ObjVente.tVente.Detail.Count - 1].GetDouble('TVE_MONTANT'));
      Obj := TObjCalculPMVR.CreateAvecTob(ObjVente.tVente, ObjVente.tOPCVM, T.GetString('TOP_GENERAL'), T.GetString('TOP_CODEOPCVM'));
      try
        Obj.CalculVente(True);
        Obj.RecupVentes(ObjVente.tVente);
      finally
        FreeAndNil(Obj);
      end;

  end

  {Vente FIFO : pour chaque ligne de la tob grille (lOpcvm) peut correspondre plusieurs enregistrements
                dans la table TROPCVM, car lOpcvm repose sur la vue TRVENTEMULTIOPCVM qui est un
                regroupement de enregistrement de la table TROPCVM par Partefeuilles, OPCVM et Généraux.
                Par rapport à une vente simple, il faut rajouter une boucle sur les lignes de lOpcvm pour
                génèrer les lignes de vente pour toutes les transactions nécessaires afin atteindre le
                nombre de parts saisies dans la grille pour chaque ligne de lOpcvm}
  else if (Multiple = vop_FIFO) or (Multiple = vop_CAMP) then begin
    {On va détruire les éventuelles Tobs concernant la clef présente dans l'objet de vente}
    SupprimerVente(T, False);
    {Gestion du nombre de parts
    15/12/06 : FQ 10373 : On passe d'un Integer à un Double}
    p := T.GetDouble('PARTAVENDRE');

//    F := TOB.Create('§§§', ObjVente.tOPCVM, -1);
    F := TOB.Create('§§§', nil, -1);
    try
      {On va chercher dans la table TROPCVM les transactions correspondant à la clef en cours}
      ChargeDetailOPCVM(F, T.GetString('TOP_PORTEFEUILLE'),
                           T.GetString('TOP_CODEOPCVM'),
                           T.GetString('TOP_GENERAL'));
      {Remise à zéro du montant éventuellement calculé précédemment}
      T.SetDouble('MONTANTVEN', 0.0);
      {On boucle sur le détail}
      for n := 0 to F.Detail.Count - 1 do begin
        {On ne traite pas les lignes sur lesquelles il n'y a rien à vendre
        15/12/06 : FQ 10373 : On passe d'un Integer à un Double}
        if (F.Detail[n].GetDouble('TOP_NBPARTVENDU') - F.Detail[n].GetDouble('TOP_NBPARTACHETE')) = 0 then Continue;

        {Création de la tob fille dans l'objet qui servira lors de la mise à jour de la table TROPCVM}
        O := TOB.Create('TROPCVM', ObjVente.tOPCVM, -1);
        O.SetString('TOP_NUMOPCVM', F.Detail[n].GetString('TOP_NUMOPCVM'));
        O.LoadDB;

        {09/05/07 : FQ 10451 : Correction du contrôle des dates de vente et d'achat. La vue TRVENTEMULTIOPCVMMS a été
                    modifié pour ne prendre en compte la plus vieille date d'achat et non plus la plus récente = >
                    le test dans le CellExit perd de son intérêts, quoiqu'il peut y avoir des gens distraits.}
        if T.GetDateTime('DATEVENTE') < O.GetDateTime('TOP_DATEACHAT') then begin
          PGIError(TraduireMemoire('La date ne peut être inférieure à celle d''achat.'), Ecran.Caption);
          T.SetDouble('MONTANTVEN', 0);
          T.SetDouble('PARTAVENDRE', 0);
          AnnuleOk := True;
          Break;
        end;

        O.AddChampSupValeur('NBPARTS', 0);
        {21/12/06 : T.GetDateTime('DATEVENTE') et F.Detail[n].GetDateTime('DATEVENTE'),
                   car dans F on a la date du Jour et dans T le contenue de la grille}
        if ObjVente.GenererVente(O, n = 0, T.GetDateTime('DATEVENTE'), p) then begin
          {On récupère le montant total de la vente dans la dernière fille (ie celle qui vient d'être
           créée) de la Tob vente de l'objet et on cumule sur la tob de la grille}
          T.SetDouble('MONTANTVEN', T.GetDouble('MONTANTVEN') + ObjVente.tVente.Detail[ObjVente.tVente.Detail.Count - 1].GetDouble('TVE_MONTANT'));
          {27/12/06 : On retranche à P les parts vendues}
          P := P - ObjVente.tVente.Detail[ObjVente.tVente.Detail.Count - 1].GetDouble('TVE_NBVENDUE');
        end;

        Obj := TObjCalculPMVR.CreateAvecTob(ObjVente.tVente, ObjVente.tOPCVM, T.GetString('TOP_GENERAL'), T.GetString('TOP_CODEOPCVM'));
        try
          Obj.CalculVente(True);
          Obj.RecupVentes(ObjVente.tVente);
        finally
          FreeAndNil(Obj);
        end;
  //      TobDebug(ObjVente.tOPCVM);

        {S'il n'y a plus de parts à vendre on sort de la boucle}
        if p <= 0 then Break;
      end;
    finally
      if Assigned(F) then FreeAndNil(F);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRVENTEOPCVM.OnAfterFormShow;
{---------------------------------------------------------------------------------------}
begin
  SendMessage(Grille.Handle, WM_KEYDOWN, VK_TAB,  0);
  SendMessage(Grille.Handle, WM_KEYDOWN, VK_LEFT, 0);
end;


initialization
  RegisterClasses([TOF_TRVENTEOPCVM]);

end.
