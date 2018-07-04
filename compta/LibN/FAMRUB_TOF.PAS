{ Unité : Gestion du paramérage des familles de rubriques
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
4.3.1.101.001    26/05/03    JP     Création de l'unité

--------------------------------------------------------------------------------------}
unit FAMRUB_TOF ;

interface

uses
  StdCtrls, Controls, Classes, forms, SysUtils, ComCtrls, HCtrls, HEnt1,
  {$IFDEF EAGLCLIENT}
   MaineAGL, UtileAGL,
  {$ELSE}
   FE_Main, db,
   {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
   UTOF,
   Graphics,
   Windows,
   Hqry,
   UTob,
   Grids,
   HTB97,
   Ent1,     // EstSpecif
   HStatus;

procedure ParametrageFamilleRubrique(Laquelle,LeLib : string ; var Fam,LibFam : string ; Budget : Boolean ; QueFListe2 : Boolean = FALSE) ;
procedure ParametragePlanRevision(Laquelle, LeLib : string ; var Fam, LibFam : string);

type
  PColor = class
    Drapeau : string ;
  end ;

  TOF_FAMRUB = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Laquelle : string ;
    LeLib    : string ;
    Budget   : Boolean ;
    OldBrush : TBrush ;
    OldPen   : TPen ;
    MaTob    : TOB;
    QueFListe2 : Boolean ;

    FBoCtxRevision : Boolean; // GCO - 22/02/2005

    procedure MajFamilleRubrique;
    procedure SelectionneFamille;
    procedure FaitRequeteMaj;
    procedure CompteElemSelectionner(Sender : TObject);
    procedure CocheDecoche          (Sender : TObject);
    procedure RempliGrid            (G : THGrid; I : Integer);
  public
    FListe1 : THGrid;
    FListe2 : THGrid;
    Nb1     : TLabel;
    Tex1    : TLabel;
    Nb2     : TLabel;
    Tex2    : TLabel;
    Qrf     : THQuery;

    procedure FlisteDblClick (Sender: TObject);
    procedure FlisteMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure FListeDrawCell (Sender: TObject; Col, Row: Longint; Rect: TRect; State: TGridDrawState);
    procedure FListeKeyDown  (Sender: TObject; var Key: Word; Shift: TShiftState);
  end ;

implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  {$ENDIF MODENT1}
  HMsgBox;

var
  _LaFam  : string ;
  _LibFam : string ;

  HMess : array [0..5] of string = (
  {00} '0;Affectation des familles de rubriques;Confirmez-vous la destruction des anciennes familles et la sauvegarde des nouvelles?;Q;YNC;Y;C;',
  {01} 'Ecriture des familles en cours',
  {02} 'Paramétrage de la rubrique',
  {03} 'Ligne sélectionnée',
  {04} 'Lignes sélectionnées',
  {05} '5');

  TChps : array [0..3, 0..2] of string =
          (
           ('CO_CODE'    , 'CO_LIBELLE' , ''),
           ('YDS_CODE'   , 'YDS_LIBELLE', ''),
           ('RB_RUBRIQUE', 'RB_LIBELLE' , 'RB_FAMILLES'),
           ('CPR_PLANREVISION', 'CPR_LIBELLE', '')
          );

{---------------------------------------------------------------------------------------}
procedure ParametrageFamilleRubrique(Laquelle,LeLib : string ; var Fam,LibFam : string ;
                                     Budget : Boolean  ; QueFListe2 : Boolean = False);
{---------------------------------------------------------------------------------------}
var
  ch ,
  B1 , B2 ,
  S1 , S2 : string;
begin
  if Budget then B1 := 'OUI'
            else B1 := 'NON';
  if QueFListe2 then B2 := 'OUI'
                else B2 := 'NON';

  {Si plusieurs familles sont passées en paramétres, elles sont séparéespar un ";",
   ce qui pose un gros problème dans le OnArgument lorsque l'on récupère les paramétres :
   On change donc le séparateur entre les familles de rubriques}
  S1 := Fam;
  while S1 <> '' do
    S2 := S2 + ReadTokenSt(S1) + '~';

  ch := Laquelle + ';' + LeLib + ';' + S2 + ';' + LibFam + ';' + B1 + ';' + B2;

  AGLLanceFiche('CP', 'CPFAMRUB', '', '', ch);

  Fam    := _LaFam ;
  LibFam := _LibFam;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/02/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure ParametragePlanRevision(Laquelle, LeLib : string ; var Fam, LibFam : string);
var S1, S2, Ch : string;
begin
  S1 := Fam;
  while S1 <> '' do
    S2 := S2 + ReadTokenSt(S1) + '~';

  Ch := Laquelle + ';' + LeLib + ';' + S2 + ';' + LibFam + ';NON;NON;OUI';

  AGLLanceFiche('CP', 'CPFAMRUB', '', '', Ch);

  Fam    := _LaFam ;
  LibFam := _LibFam;
end;

////////////////////////////////////////////////////////////////////////////////
{---------------------------------------------------------------------------------------}
procedure TOF_FAMRUB.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
var
  SQL, ch : string;
  Ind : Integer;
  Q : TQuery;
  lStWhere : string;
begin
  inherited ;

  {Récupération des paramètres de création}
  Laquelle   := ReadTokenSt(S);
  LeLib      := ReadTokenSt(S);
  ch         := ReadTokenSt(S);
  _LibFam    := ReadTokenSt(S);
  Budget     := ReadTokenSt(S) = 'OUI';
  QueFliste2 := ReadTokenSt(S) = 'OUI';
  FBoCtxRevision := ReadTokenSt(S) = 'OUI';

  {On remet le ";" comme séparateur des familles de rubriques après l'avoir enlevé
   dans ParametrageFamilleRubrique}
  _LaFam := '';
  while not (ch = '') do
    _LaFam := _LaFam + ReadTokenPipe(ch, '~') + ';';

  OldBrush := TBrush.Create ;
  OldPen   := TPen.Create ;

  {Pointeurs sur les deux grilles de la fiche}
  FListe1 := THGrid(GetControl('FLISTE1',True));
  FListe2 := THGrid(GetControl('FLISTE2',True));

  Nb1  := TLabel(GetControl('NB1' , True));
  Tex1 := TLabel(GetControl('TEX1', True));
  Nb2  := TLabel(GetControl('NB2' , True));
  Tex2 := TLabel(GetControl('TEX2', True));

  TToolBarButton97(GetControl('BDELETE' , True)).Visible := False;
  TToolBarButton97(GetControl('BINSERT' , True)).Visible := False;
  TToolBarButton97(GetControl('BDEFAIRE', True)).Visible := False;
  TToolBarButton97(GetControl('BIMPRIMER',True)).Visible := False;

  {Affectation des évènements de la fiche}
  Fliste1.OnDblClick  := FlisteDblClick;
  Fliste1.OnMouseDown := FlisteMouseDown;
  Fliste1.OnDrawCell  := FListeDrawCell;
  Fliste1.OnKeyDown   := FListeKeyDown;
  Fliste2.OnDblClick  := FlisteDblClick;
  Fliste2.OnMouseDown := FlisteMouseDown;
  Fliste2.OnDrawCell  := FListeDrawCell;
  Fliste2.OnKeyDown   := FListeKeyDown;

  if FBoCtxRevision then
  begin
    if EstSpecif('51502') and (ctxStandard in V_PGI.PGIContexte) then
      lStWhere := 'CPR_PREDEFINI = "CEG"'
    else
    begin
      { CA - 04/10/2005 - En mode PCL, on restreint la liste aux standards. En mode PGE, on exclu seulement les CEG }
      if ctxPCL in V_PGI.PGIContexte then
        lStwhere := 'CPR_PREDEFINI = "STD"'
      else
        lStwhere := 'CPR_PREDEFINI <> "CEG"'
    end;

    Sql := 'SELECT CPR_PLANREVISION, CPR_LIBELLE FROM CREVPLAN WHERE ' +
            lStWhere + ' ORDER BY CPR_PLANREVISION';

    Ind := 3;
  end
  else
  begin
    if Budget then begin
      Sql := 'SELECT CO_CODE,CO_LIBELLE FROM COMMUN WHERE CO_TYPE="RBB" ORDER BY CO_CODE ';
      Ind := 0;
    end
    else begin
      Sql := 'SELECT YDS_CODE,YDS_LIBELLE FROM CHOIXDPSTD WHERE YDS_TYPE="RBF" ORDER BY YDS_CODE ';
      Ind := 1;
    end;
  end;

  Q := OpenSql(Sql, False);
  if MaTob = nil then
    MaTob := TOB.Create('', nil, -1);
  MaTob.LoadDetailDB('CHOIXDPSTD','','',Q, False);
  Ferme(Q);

  RempliGrid(FListe2, Ind);

  if (Laquelle = '') and (not QueFListe2) then begin
    FaitRequeteMaj;
    RempliGrid(FListe1, 2);
    FListe1.SetFocus;
    Nb2 .Left := 376;
    Tex2.Left := 390;
    Nb2 .Caption := '0';
    Tex2.Caption := HMess[3];
    Nb1 .Caption := '0';
    Tex1.Caption := HMess[3];
    Ecran.HelpContext := 7775200;
  end
  else begin
    Fliste1.Visible := False;
    Ecran.Caption := HMess[2] + ' ' + Laquelle + ' ' + LeLib;
    Ecran.Width := 260;
    FListe2.Align := alClient;
    Nb1 .Visible := False;
    Tex1.Visible := False;
    Nb2 .Left := 6;
    Tex2.Left := 20;
    Nb2 .Caption := '0';
    Tex2.Caption := HMess[3];
    SelectionneFamille;
    FListe2.SetFocus;
    Ecran.HelpContext := 7775000 ;
  end;
  UpdateCaption(Ecran) ;

  Application.ProcessMessages;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_FAMRUB.FlisteMouseDown(Sender: TObject; Button: TMouseButton;
                                    Shift: TShiftState; X, Y: Integer);
{---------------------------------------------------------------------------------------}
begin
  if ((ssCtrl in Shift) and (Button = mbLeft)) then
    CocheDecoche(Sender) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_FAMRUB.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
{---------------------------------------------------------------------------------------}
begin
  if (Key = VK_SPACE) then
    CocheDecoche(Sender) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_FAMRUB.FListeDrawCell(Sender: TObject; Col, Row: Longint; Rect: TRect;
                                   State: TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  Text   : array[0..255] of Char;
  F      : TAlignment ;
  Grille : THgrid ;
begin

  Grille := THgrid(Sender) ;
  OldBrush.assign(Grille.Canvas.Brush) ;
  OldPen.assign(Grille.Canvas.Pen) ;
  StrPCopy(Text,Grille.Cells[Col,Row]);
  Grille.Canvas.Font.Style:=Grille.Canvas.Font.Style-[fsItalic] ;

  if (gdFixed in State) then begin
    Grille.Canvas.Brush.Color:=Grille.FixedColor ;
    Grille.Canvas.Font.Color:=Grille.Font.Color ;
    F:=taCenter
  end
  else begin
    F := taLeftJustify ;
    if (gdSelected in State) then begin
      if PColor(Grille.Objects[0,Row]).Drapeau='*' then begin
         Grille.Canvas.Brush.Color:=clInactiveCaption ;
         Grille.Canvas.Font.Color:=clHighlightText ;
         Grille.Canvas.Font.Style:=Grille.Canvas.Font.Style+[fsItalic] ;
       end
       else begin
         Grille.Canvas.Brush.Color:=clHighlight ;
         Grille.Canvas.Font.Color:=clHighlightText ;
       end;
     end
     else begin
       if PColor(Grille.Objects[0,Row]).Drapeau='*' then begin
         Grille.Canvas.Brush.Color:=clInactiveCaption ;
         Grille.Canvas.Font.Color:=cl3DLight ;
       end
       else begin
         Grille.Canvas.Brush.Color:=clWindow ;
         Grille.Canvas.Font.Color:=clWindowText ;
       end;
    end;
  end;

  case F of
    taRightJustify : ExtTextOut(Grille.Canvas.Handle, Rect.Right - Grille.Canvas.TextWidth(Grille.Cells[Col,Row])-3,
                               Rect.Top + 2,ETO_OPAQUE or ETO_CLIPPED, @Rect, Text, StrLen(Text), nil) ;
    taCenter       : ExtTextOut(Grille.Canvas.Handle,Rect.Left + ((Rect.Right-Rect.Left-Grille.canvas.TextWidth(Grille.Cells[Col,Row])) div 2),
                         Rect.Top + 2, ETO_OPAQUE or ETO_CLIPPED, @Rect, Text, StrLen(Text), nil) ;
    else
      ExtTextOut(Grille.Canvas.Handle, Rect.Left + 2, Rect.Top + 2, ETO_OPAQUE or ETO_CLIPPED,
                 @Rect, Text, StrLen(Text), nil) ;
  end;

  if ((gdfixed in State) and Grille.Ctl3D) then begin
    DrawEdge(Grille.Canvas.Handle, Rect, BDR_RAISEDINNER, BF_BOTTOMRIGHT);
    DrawEdge(Grille.Canvas.Handle, Rect, BDR_RAISEDINNER, BF_TOPLEFT);
  end;

  Grille.Canvas.Brush.Assign(OldBrush) ;
  Grille.Canvas.Pen.Assign(OldPen) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_FAMRUB.FlisteDblClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  CocheDecoche(Sender);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_FAMRUB.OnNew ;
{---------------------------------------------------------------------------------------}
begin
  inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_FAMRUB.OnDelete ;
{---------------------------------------------------------------------------------------}
begin
  inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_FAMRUB.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  i : Integer ;
begin
  inherited ;
  if (Laquelle = '') and (not QueFListe2) then begin
    if HShowMessage(HMess[0],'','') = mrYes then begin
      MajFamilleRubrique;
      FaitRequeteMaj;
      RempliGrid(FListe1, 2);
      CompteElemSelectionner(FListe1);

      for i := 1 to FListe2.RowCount - 1 do
        PColor(FListe2.Objects[0,i]).Drapeau := '';

      FListe2.Row := 1;
      FListe2.Invalidate;
      CompteElemSelectionner(FListe2) ;
    end;
  end
  else begin
    _LaFam := '' ;
    for i := 1 to Fliste2.RowCount - 1 do
      if PColor(FListe2.Objects[0,i]).Drapeau = '*' then begin
        _LaFam  := _LaFam  + Fliste2.Cells[0,i] + ';' ;
        _LibFam := _LibFam + Fliste2.Cells[1,i] + ';' ;
      end;
  end;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_FAMRUB.OnLoad ;
{---------------------------------------------------------------------------------------}
begin
  inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_FAMRUB.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  FListe1.VidePile(True);
  FListe2.VidePile(True);
  OldBrush.Free;
  OldPen.Free;
  inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_FAMRUB.OnCancel () ;
{---------------------------------------------------------------------------------------}
begin
  inherited ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_FAMRUB.SelectionneFamille ;
{---------------------------------------------------------------------------------------}
var
  St  : string;
  St1 : string;
  i   : Byte;
begin
  if _LaFam='' then Exit ;
  St := _Lafam;
  while St <> '' do begin
    St1 := ReadTokenSt(St);
    for i := 1 to Fliste2.RowCount - 1 do
      if Fliste2.Cells[0,i] = St1 then begin
        PColor(FListe2.Objects[0,i]).Drapeau := '*';
        Break ;
      end;
  end;
  CompteElemSelectionner(FListe2) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_FAMRUB.RempliGrid(G : THGrid; I : Integer);
{---------------------------------------------------------------------------------------}
var
  n  : Integer ;
  X  : PColor ;
  St : string ;
  ch : string ;
begin
  G.VidePile(True);

  for n := 0 to MaTob.Detail.Count - 1 do begin
    X := PColor.Create;
    X.Drapeau := '';
    {A mon grand désespoir, getNomChamp et GetValeur ne semblent pas fonctionner,
    d'où l'utilisation d'un tableau indicé pour récupérer les noms de champs}
    St := MaTob.Detail[n].GetValue(TChps[I,0]);
    ch := MaTob.Detail[n].GetValue(TChps[I,1]);

    if Budget and (UpperCase(G.Name) = 'FLISTE2') then begin
           if St = 'ANA' then St := 'CBS'
      else if St = 'A/G' then St := 'S/G'
      else if St = 'GEN' then St := 'CBG'
      else if St = 'G/A' then St := 'G/S';
    end;

    G.RowCount := G.RowCount + 1;
    G.Cells  [0,n+1] := St;
    G.Objects[0,n+1] := X;
    G.Cells  [1,n+1] := ch;

    if UpperCase(G.Name) = 'FLISTE1' then
      G.Cells[2,n+1] := MaTob.Detail[n].GetValue(TChps[I,2]);
  end; {for n}

  if G.RowCount > 2 then
    G.RowCount:=G.RowCount - 1;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_FAMRUB.CocheDecoche(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Grille : THGrid;
begin
  if Sender = nil then Exit ;
  Grille := THGrid(Sender) ;

  if PColor(Grille.Objects[0,Grille.Row]).Drapeau = '*' then
    PColor(Grille.Objects[0,Grille.Row]).Drapeau := ''
  else
    PColor(Grille.Objects[0,Grille.Row]).Drapeau := '*';

  Grille.Invalidate ;
  CompteElemSelectionner(Grille);

  Grille.Refresh;
  Nb1   .Refresh;
  Tex1  .Refresh;
  Nb2   .Refresh;
  Tex2  .Refresh;

  Application.ProcessMessages;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_FAMRUB.MajFamilleRubrique ;
{---------------------------------------------------------------------------------------}
var
  i : Integer;
  j : Integer;
  St  : string ;
  Sql : string ;
begin
  j := 0;

  for i := 1 to Fliste1.RowCount - 1 do
    if PColor(Fliste1.Objects[0,i]).Drapeau = '*' then Inc(j) ;

  if j > 0 then begin
    for i:=1 to Fliste2.RowCount-1 do
      if PColor(Fliste2.Objects[0,i]).Drapeau = '*' then
        St := St+Fliste2.Cells[0,i] + ';';

     BeginTrans;
     InitMove(j, HMess[1]);

     for i := 1 to Fliste1.RowCount - 1 do
       if PColor(Fliste1.Objects[0, i]).Drapeau = '*' then begin
         Sql := 'Update RUBRIQUE Set RB_FAMILLES="' + St + '" Where RB_RUBRIQUE="'
                + FListe1.Cells[0, i] + '"';
         ExecuteSQL(Sql);
         MoveCur(False) ;
       end; {if PColor(}

    FiniMove ;
    CommitTrans ;
  end; {if j > 0 }
end;

{---------------------------------------------------------------------------------------}
Procedure TOF_FAMRUB.CompteElemSelectionner(Sender : TObject) ;
{---------------------------------------------------------------------------------------}
var
  i : Integer;
  j : Integer;
  C : Char;
begin
  j := 0;
  C := THGrid(Sender).Name[Length(THGrid(Sender).Name)] ;

  for i := 1 to THGrid(Sender).RowCount - 1 do
    if PColor(THGrid(Sender).Objects[0,i]).Drapeau = '*' then Inc(j) ;

  case j of
    0, 1: begin
            if C='1' then begin
              Nb1 .Caption := IntTostr(j);
              Tex1.Caption := HMess[3] ;
            end
            else begin
              Nb2 .Caption := IntTostr(j);
              Tex2.Caption := HMess[3] ;
            end;
          end;

    else begin
      if C='1' then begin
        Nb1 .Caption := IntTostr(j);
        Tex1.Caption := HMess[4] ;
      end
      else begin
        Nb2 .Caption := IntTostr(j);
        Tex2.Caption := HMess[4] ;
      end; {if else}
    end; {case else}
  end; {case}
end;

{---------------------------------------------------------------------------------------}
Procedure TOF_FAMRUB.FaitRequeteMaj ;
{---------------------------------------------------------------------------------------}
var
  Sql : string;
  Q   : TQuery;
begin
  if Budget then Sql := 'SELECT RB_RUBRIQUE,RB_LIBELLE,RB_FAMILLES FROM RUBRIQUE WHERE RB_NATRUB="BUD" ORDER BY RB_RUBRIQUE'
            else Sql := 'SELECT RB_RUBRIQUE,RB_LIBELLE,RB_FAMILLES FROM RUBRIQUE WHERE RB_NATRUB<>"BUD" ORDER BY RB_RUBRIQUE';

  Q := OpenSql(Sql, False);
  if MaTob = nil then
    MaTob := TOB.Create('', nil, -1);
  MaTob.LoadDetailDB('RUBRIQUE','','',Q, False);
  Ferme(Q);
end;

Initialization
  registerclasses ( [ TOF_FAMRUB ] ) ;

end.
