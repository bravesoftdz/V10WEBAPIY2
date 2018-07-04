{-------------------------------------------------------------------------------------
    Version   |   Date  | Qui | Commentaires
--------------------------------------------------------------------------------------
 6.50.001.001  05/01/05   VL    Création de l'unité (TOM de la TABLE : REPARTBU)
 6.50.001.006  07/07/05   JP    FQ 16006 : Impossible d'enregistrer les modifications
 6.50.001.015  25/08/05   JP    FQ 16007 : Ajout d'un bouton de duplication
--------------------------------------------------------------------------------------}
unit REPARTBU_TOM ;

interface

uses StdCtrls, Controls, Classes, Windows,
  {$IFDEF EAGLCLIENT}
    eFichList, MainEagl,
  {$ELSE}
    db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    FichList, FE_Main,
  {$ENDIF}
    forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, Ent1, SAISUTIL,
    utilPGI,  // _Blocage
    UTOM, UTob;

Procedure ParamRepartBUDGET ( Quel : String ; ForceCreat : boolean ) ;

Type
  TOM_REPARTBU = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  private
    GR : THGrid;
    TotalTaux : THNumEdit;
    BR_BUDJAL, PEREXODEB, PEREXOFIN : THValComboBox;
    PerDeb,PerFin : TDateTime;
    NbPeriodes : Integer;
    {JP 24/08/05 : FQ 16007 : Pour ne pas fermer la fiche sur le BValider.Click appelé
                   depuis la duplication}
    PeutSortir  : Boolean;
    FCloseQuery : TCloseQueryEvent;

    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BR_BUDJALChange(Sender : TObject);

    {JP 24/08/05 : FQ 16007 : Duplication de la clef pour contourner le fait que le budget
                   appartienne à la clef primaire de la table}
    procedure BDupliquerClick  (Sender : TObject);
    procedure FormCanCloseQuery(Sender : TObject; var CanClose : Boolean);

    procedure GRCellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GRCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GRRowExit(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure GRSetEditText(Sender: TObject; ACol, ARow: Longint; const Value: HString);

    procedure EcritLesTaux;
    procedure SelectNbTaux;
    procedure ChercheNbMois;
    procedure MontreTotal;
    function  TotalTauxOK : Boolean;
    procedure RempliTaux;

    {JP 25/08/05 : FQ 16007 : Processus de suplication de la clef courante}
    procedure DuplicationClef;
  end;

implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  HTB97;

Procedure ParamRepartBUDGET ( Quel : String ; ForceCreat : boolean ) ;
var
  szArg : String;
begin
  if _Blocage(['nrCloture',''],False,'') then Exit;
  if ForceCreat then szArg := Quel+';ACTION=CREATION'
                else szArg := Quel;
  AGLLanceFiche('CP', 'CPREPARTBU', '', '', szArg);
end;

procedure TOM_REPARTBU.OnNewRecord ;
begin
  Inherited ;
  GR.VidePile(True);
  GR.Enabled:=False;
  TotalTaux.Visible := False;
  SetControlVisible('TTotalTaux', False);
  SetField('BR_TAUX', '');
  SelectNbTaux;
end ;

procedure TOM_REPARTBU.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_REPARTBU.OnUpdateRecord ;
begin
  inherited;
  {Vérification de la saisie d'un code budget}
  if GetField('BR_BUDJAL')='' then begin
    LastErrorMsg := TraduireMemoire('Budget non renseigné. La saisie est obligatoire.');
    LastError := 1;
    Exit;
  end;

  {Vérification du total des période}
  if not TotalTauxOK then begin
    LastErrorMsg := TraduireMemoire('Le total des taux ne fait pas 100%.');
    LastError := 1;
    Exit;
  end;

  {Remplissage du champ memo avec les données de la grille des périodes}
  EcritLesTaux;

  {Si l'on est en création on s'assure}
  if (DS.State = dsInsert) and
     ExisteSQL('SELECT BR_REPARPERIODE FROM REPARTBU WHERE BR_REPARPERIODE = "' + GetField('BR_REPARPERIODE') + '"') then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Cette fiche existe déjà. Vous devez la modifier.');
  end;
end;

procedure TOM_REPARTBU.OnAfterUpdateRecord;
begin
  Inherited ;
end ;

procedure TOM_REPARTBU.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_REPARTBU.OnLoadRecord ;
begin
  Inherited ;
  SelectNbTaux;
  MontreTotal;
end ;

procedure TOM_REPARTBU.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_REPARTBU.OnArgument ( S: String ) ;
var
  Quel : String;
  Bt   : TToolbarButton97; {JP 24/08/05 : FQ 16007}
begin
  Inherited ;

  Quel := ReadTokenST(S);

  Ecran.OnKeyDown := FormKeyDown;
  GR := THGrid(GetControl('GR', True));
  GR.OnCellEnter := GRCellEnter;
  GR.OnCellExit := GRCellExit;
  GR.OnRowExit := GRRowExit;
  GR.OnSetEditText := GRSetEditText;
  BR_BUDJAL := THValComboBox(GetControl('BR_BUDJAL', True));
  BR_BUDJAL.OnChange := BR_BUDJALChange;
  TotalTaux := THNumEdit(GetControl('TOTALTAUX', True));
  PEREXODEB := THValComboBox(GetControl('PEREXODEB', True));
  PEREXOFIN := THValComboBox(GetControl('PEREXOFIN', True));

  {JP 24/08/05 : FQ 16007 : Duplication de la clef courante}
  Bt := TToolbarButton97(GetControl('BDUPLIQUER'));
  if Assigned(Bt) then
    Bt.OnClick := BDupliquerClick;

  {JP 24/08/05 : FQ 16007 : Gestion du CloseQuery}
  PeutSortir := True;
  FCloseQuery := TFFicheListe(Ecran).FormCloseQuery;
  TFFicheListe(Ecran).OnCloseQuery := FormCanCloseQuery;

  ChangeMask(TotalTaux, ADecimP, '');
//  ChangeSizeMemo(TaBR_TAUX); // Qu'est-ce que cela fait ??
end ;

procedure TOM_REPARTBU.OnClose ;
begin
  Inherited ;
  GR.VidePile(True);
end ;

procedure TOM_REPARTBU.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_REPARTBU.BR_BUDJALChange(Sender: TObject);
begin
  GR.Enabled:=True;
  TotalTaux.Visible := True;
  SetControlVisible('TTotalTaux', True);
  SelectNbTaux;
  MontreTotal;
end;

procedure TOM_REPARTBU.EcritLesTaux;
var
  i : Byte ;
  St : String;
BEGIN
  St:='';
  for i:=1 to NbPeriodes do
    St:=St+GR.Cells[1,i]+';';
  SetField('BR_TAUX', St);
end;

procedure TOM_REPARTBU.SelectNbTaux;
begin
  ChercheNbMois;
  RempliTaux;
  if DS.State in [dsBrowse] then MontreTotal;
end;

procedure TOM_REPARTBU.ChercheNbMois;
var
  Q : TQuery;
  ExoDeb,ExoFin : String;
  Cur : Integer;
  DD  : TDateTime;
  St  : String;
BEGIN
  if BR_BUDJAL.Value='' then Exit;
  Q := OpenSQL('SELECT BJ_EXODEB, BJ_EXOFIN, BJ_PERDEB, BJ_PERFIN FROM BUDJAL WHERE BJ_BUDJAL="'+BR_BUDJAL.Value+'"', True);
  if not Q.EOF then begin
    ExoDeb := Q.Fields[0].AsString;
    ExoFin:=Q.Fields[1].AsString;
    if (ExoDeb='') Or (ExoFin='') then begin
      Ferme(Q);
      Exit;
    end;
    ListePeriode(ExoDeb, PEREXODEB.Items, PEREXODEB.Values, True);
    ListePeriode(ExoFin, PEREXOFIN.Items, PEREXOFIN.Values, True);
    PerDeb := Q.Fields[2].AsDateTime;
    PerFin := Q.Fields[3].AsDateTime;
    PEREXODEB.Value := DateToStr(DebutDeMois(PerDeb));
    PEREXOFIN.Value := DateToStr(DebutDeMois(PerFin));
  end;
  Ferme(Q);
  NbPeriodes := 0;
  if ExoDeb=ExoFin then begin
    Cur := PEREXODEB.ItemIndex;
    While Cur<=PEREXOFIN.ItemIndex do begin
      DD := StrToDate(PEREXODEB.Values[Cur]);
      St := FormatDateTime('mmmm yy',DD);
      GR.Cells[0,Cur-PEREXODEB.ItemIndex+1] := FirstMajuscule(St);
      Inc(Cur);
      Inc(NbPeriodes);
    end;
    end
  else begin
    for Cur:=PEREXODEB.ItemIndex to PEREXODEB.Items.Count-1 do begin
      DD := StrToDate(PEREXODEB.Values[Cur]);
      St := FormatDateTime('mmmm yy',DD);
      GR.Cells[0,Cur-PEREXODEB.ItemIndex+1]:=FirstMajuscule(St);
      Inc(NbPeriodes);
    end;
    for Cur:=0 to PEREXOFIN.ItemIndex do begin
      DD := StrToDate(PEREXOFIN.Values[Cur]);
      St := FormatDateTime('mmmm yy',DD);
      GR.Cells[0,Cur+PEREXODEB.Items.Count-PEREXODEB.ItemIndex+1] := FirstMajuscule(St);
      Inc(NbPeriodes);
    end;
  end;
  GR.RowCount := 1+NbPeriodes;
end;

procedure TOM_REPARTBU.MontreTotal;
var
  TTaux : Double;
  i : Byte;
begin
  TTaux := 0;
  for i:=1 to NbPeriodes do
    TTaux := TTaux+Valeur(GR.Cells[1,i]);
  TotalTaux.Value:=TTaux;
end;

function TOM_REPARTBU.TotalTauxOK: Boolean;
begin
  Result := Arrondi(TotalTaux.Value-100.0,ADecimP)=0;
end;

procedure TOM_REPARTBU.RempliTaux;
var
  St,StC : String;
  ii     : integer;
  X      : Double;
begin
  for ii:=1 to GR.RowCount-1 do
    GR.Cells[1,ii]:='';
  St := Trim(GetField('BR_TAUX'));
  ii:=0;
  repeat
    StC := ReadTokenSt(St);
    inc(ii);
    if StC<>'' then begin
      X := Valeur(StC);
      GR.Cells[1,ii] := StrFMontant(X, 15, ADecimP, '', True);
    end;
  until St='';
end;

procedure TOM_REPARTBU.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if ((Key=VK_RETURN) and (Shift=[]) and (GR.Focused)) then Key:=VK_TAB;
end;

procedure TOM_REPARTBU.GRCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  {JP 07/07/05 : FQ 16006 : inutile, cela empêche la validation et c'est fait dans GRSetEditText
  DS.Edit;}
end;

procedure TOM_REPARTBU.GRCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  X : double;
begin
  X := Valeur(GR.Cells[ACol,ARow]);
  if X=0 then GR.Cells[ACol,Arow] := ''
         else GR.Cells[ACol,ARow] := StrFMontant(X, 15, ADecimP, '', True);
end;

procedure TOM_REPARTBU.GRRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  MontreTotal;
end;

procedure TOM_REPARTBU.GRSetEditText(Sender: TObject; ACol, ARow: Integer; const Value: HString);
begin
  // Forcer la modif
  if DS.State=dsBrowse then begin
    DS.Edit;
    SetField('BR_Societe', '');
    SetField('BR_Societe', V_PGI.CodeSociete);
  end;
end;

{JP 24/08/05 : FQ 16007 : Duplication de la clef pour contourner le fait que le budget
               appartienne à la clef primaire de la table : en 2 / 3, ce n'est pas bloquant
               et cela permet de simuler une duplication
{---------------------------------------------------------------------------------------}
procedure TOM_REPARTBU.BDupliquerClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {On n'autorise pas la duplication si on est en mode insertion}
  if DS.State = dsInsert then begin
    HShowMessage('1;' + Ecran.Caption + ';Vous êtes en création. Veuillez enregistrez vos modification'#13 +
                 'avant de dupliquer la clef courante.;W;O;O;O;', '', '');
    Exit;
  end;

  {Pour ne pas fermer la fiche (utiliser dans le FormCanCloseQuery}
  PeutSortir := False;

  try
    {Le Speed button ne prenant pas le focus, on force la sortie du contrôle courant
     pour forcé, si besoin, l'état du Datasource}
    TFFicheListe(Ecran).FListe.SetFocus;

    if HShowMessage('0;' + Ecran.Caption + ';Voulez-vous dupliquer les données de la clef courante ?;Q;YNC;Y;C;', '', '') = mrYes then
      DuplicationClef;

  finally
    PeutSortir := True;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_REPARTBU.FormCanCloseQuery(Sender : TObject; var CanClose : Boolean);
{---------------------------------------------------------------------------------------}
begin
  CanClose := PeutSortir;
  if Assigned(FCloseQuery) then FCloseQuery(Sender, CanClose);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_REPARTBU.DuplicationClef;
{---------------------------------------------------------------------------------------}
var
  Grille : string;
  Libel  : string;
begin
   {On mémorise l'enregistrement courant}
   Grille := GetField('BR_TAUX');
   Libel  := GetField('BR_LIBELLE');
   {Insertion d'un nouvel enregistrement}
   TFFicheListe(Ecran).Binsert.Click;
   {Initialisation avec les valeurs mémorisées}
   SetField('BR_TAUX', Grille);
   SetField('BR_LIBELLE', Libel);
   {Chargement de la grille}
   OnLoadRecord;
end;

initialization
  RegisterClasses ( [ TOM_REPARTBU ] ) ;

end.
