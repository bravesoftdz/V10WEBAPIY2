unit VentAna ;

interface

Uses HEnt1,
     HCtrls,
     UTOB,
     Ent1,
{$IFNDEF EAGLSERVER}
     LookUp,
     AGLInit,
{$ENDIF EAGLSERVER}
     Controls,
     ComCtrls,
     StdCtrls,
     ExtCtrls,
     SysUtils,
     Dialogs,
     SaisUtil,
     UtilPGI,
     UtilSais,
     EntGC,
     Classes,
     HMsgBox,
     SaisComm,
     HSysMenu,
     Grids,
     Buttons,
     HPanel,
{$IFDEF EAGLCLIENT}
{$ELSE}
     Hcompte,
{$IFNDEF EAGLSERVER}
     FE_Main,
     EdtRDoc,
{$ENDIF EAGLSERVER}
     DB,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     ed_tools,
{$ENDIF}
     Forms,
     Windows,
     HTB97,
     Mask,
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ENDIF MODENT1}
     messages, TntStdCtrls, TntGrids, TntButtons;

{$IFDEF GG}      //Pour Gérard Piot Version 573
Procedure YYVentilAna ( TOBL,TOBAna : TOB ; Action : TActionFiche; LaReferenceAna : String='' ) ;
{$ELSE}
function YYVentilAna ( TOBL,TOBAna : TOB ; Action : TActionFiche; Ind : integer ; LaReferenceAna : String=''; FromGC : boolean=false ) : boolean;
{$ENDIF}


type
  TFVentAna = class(TForm)
    HPanel1: TPanel;
    Pages: TPageControl;
    TFType: THLabel;
    FType: THValComboBox;
    MsgBox: THMsgBox;
    HMTrad: THSystemMenu;
    Panel6: TPanel;
    OKBtn: THBitBtn;
    BFermer: THBitBtn;
    HelpBtn: THBitBtn;
    Panel1: TPanel;
    HLabel1: THLabel;
    FTotVal: THNumEdit;
    Label1: TLabel;
    TS1: TTabSheet;
    FListe1: THGrid;
    TS2: TTabSheet;
    FListe2: THGrid;
    Label2: TLabel;
    TCache: THLabel;
    BSolde: THBitBtn;
    TS3: TTabSheet;
    FListe3: THGrid;
    procedure FormShow(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
{$IFNDEF EAGLSERVER}
    procedure FListe1DblClick(Sender: TObject);
{$ENDIF EAGLSERVER}
    procedure FListe1CellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure FListe1CellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure FTypeClick(Sender: TObject);
    procedure BFermerClick(Sender: TObject);
    procedure FListe1RowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure HelpBtnClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure PagesChange(Sender: TObject);
    procedure BSoldeClick(Sender: TObject);
    procedure FListe1ColExit(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure FListe1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private { Déclarations private }
    CodeSectOk,FClosing : Boolean ;
    LaReference,StCellCur : String;
    FromGC : boolean;
    procedure Recalc (F : THGrid);
    procedure ChargeVentil ;
{$IFNDEF EAGLSERVER}
    Function  ChercheSect( GS : THGrid ; C,L : integer ) : byte ;
    Procedure InitTaux(F : THGrid ; ARow : Integer) ;
{$ENDIF EAGLSERVER}
    Procedure NumeroteLigne(Sender : TObject) ;
    Function  SectionExiste(Sender :  TObject ; ARow : Integer) : Boolean ;
    procedure InitTitre ;
    procedure DeleteLigne;
  public  { Déclarations public }
    TOBAna,TOBL : TOB ;
    Action      : TActionFiche ;
    Ind         : integer ;
    retOk : boolean;
  end;

implementation
uses UFonctionsCBP
{$IFNDEF EAGLCLIENT}
{$IFNDEF EAGLSERVER}
{Section,},CPSECTION_TOM,PrintDBG
{$ENDIF EAGLSERVER}
{$ENDIF}
;

{$R *.DFM}

{$IFDEF GG}
Procedure YYVentilAna ( TOBL,TOBAna : TOB ; Action : TActionFiche; LaReferenceAna : String='' ) ;
{$ELSE}
function YYVentilAna ( TOBL,TOBAna : TOB ; Action : TActionFiche; Ind : integer ; LaReferenceAna : String=''; FromGC : boolean=false ) : boolean;
{$ENDIF}
Var FVentil: TFVentAna ;
BEGIN
result := false;
FVentil:=TFVentAna.Create(Application) ;
FVentil.TOBAna:=TOBAna ;
FVentil.TOBL:=TOBL ;
FVentil.Action:=Action ;
FVentil.FromGC := FromGC;
{ GC_DBR_Pour supprimer la directive de compilation en GC }
{$IFDEF GCGC}
FVentil.Ind:=0 ;
{$ELSE GCGC}
{$IFDEF VRELEASE}
FVentil.Ind:=0 ;
{$ELSE VRELEASE}
FVentil.Ind:=Ind ;
{$ENDIF !VRELEASE}
{$ENDIF !GCGC}
FVentil.LaReference := LaReferenceAna;
fventil.retOk := false;
try
  FVentil.ShowModal ;
  finally
  if Fventil.retOk then result := true;
  FVentil.Free ;
  end;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFVentAna.ChargeVentil ;
Var i : integer ;
    TOBA : TOB ;
    Section : String ;
    FListe : THGrid ;
BEGIN
FListe1.VidePile(False) ; FListe2.VidePile(False) ; FListe3.VidePile(False) ;
for i:=0 to TOBAna.Detail.Count-1 do
    BEGIN
    TOBA:=TOBAna.Detail[i] ;
    if TOBA.GetValue('YVA_AXE')='A1' then FListe:=FListe1 else
     if TOBA.GetValue('YVA_AXE')='A2' then FListe:=FListe2 else
        FListe:=FListe3 ;
    FListe.Cells[0,FListe.RowCount-1]:=IntToStr(TOBA.GetValue('YVA_NUMVENTIL')) ;
    Section:=TOBA.GetValue('YVA_SECTION') ;
    FListe.Cells[1,FListe.RowCount-1]:=Section ;
    if VH^.LiaisonY2ViaShare then
    begin
      if FListe=FListe1 then FListe.Cells[2,FListe.RowCount-1]:=RechDom('TZCSECTION',Section,False)
                        else FListe.Cells[2,FListe.RowCount-1]:=RechDom('TZCSECTION2',Section,False) ;
    end else
    begin
      if FListe=FListe1 then FListe.Cells[2,FListe.RowCount-1]:=RechDom('TZSECTION',Section,False)
                        else FListe.Cells[2,FListe.RowCount-1]:=RechDom('TZSECTION2',Section,False) ;
    end;

    FListe.Cells[3,FListe.RowCount-1]:=StrFMontant(TOBA.GetValue('YVA_POURCENTAGE'),15,4,'',TRUE) ;
    FListe.RowCount:=FListe.RowCount+1 ;
    END ;
END ;

procedure TFVentAna.Recalc (F : THGrid);
Var i : Integer ;
    V : Double ;
BEGIN
V:=0 ;
For i:=1 to F.RowCount-1 do
    BEGIN
    V:=V+Valeur(F.Cells[3,i]) ;
    END ;
if F.Cells[1,F.RowCount-1]<>'' then F.RowCount:=F.RowCount+1 ;
NumeroteLigne(F) ;
FTotVal.Value:=V ;
END ;

{$IFNDEF EAGLSERVER}
Procedure TFVentAna.InitTaux(F : THGrid ; ARow : Integer) ;
BEGIN
if F.Cells[3,ARow]='' then F.Cells[3,ARow]:=StrFMontant(Valeur(F.Cells[3,ARow]),15,4,'',TRUE) ;
END ;
{$ENDIF EAGLSERVER}

procedure TFVentAna.InitTitre ;
BEGIN
if TOBL=Nil then Exit ;
if TOBL.NomTable='PIECE' then
   BEGIN
   if Ind<=0 then Caption:=MsgBox.Mess[5] else Caption:=MsgBox.Mess[8] ;
   END else
if TOBL.NomTable='LIGNE' then
   BEGIN
   if Ind<=0 then Caption:=MsgBox.Mess[6] else Caption:=MsgBox.Mess[9] ;
   END else
 if TOBL.NomTable='PAIEENCOURS' then Caption:=MsgBox.Mess[7] else
     Exit ;
UpdateCaption(Self) ;
END ;

procedure TFVentAna.FormShow(Sender: TObject);
begin
{$IFNDEF PAIEGRH}
if Not VH_GC.GCVentAxe3 then TS3.TabVisible:=False ;
if Not VH_GC.GCVentAxe2 then TS2.TabVisible:=False ;
{$ENDIF}
ChargeVentil ; CodeSectOk:=False ; Recalc(FListe1) ;
if Action=taConsult then
   BEGIN
   FListe1.Options:=FListe1.Options-[goEditing,goAlwaysShowEditor] ;
   FListe1.Options:=FListe1.Options-[goRowSelect] ;
   FListe2.Options:=FListe2.Options-[goEditing,goAlwaysShowEditor] ;
   FListe2.Options:=FListe2.Options-[goRowSelect] ;
   FListe3.Options:=FListe3.Options-[goEditing,goAlwaysShowEditor] ;
   FListe3.Options:=FListe3.Options-[goRowSelect] ;
   FType.Enabled:=False ;
   END ;
InitTitre ;
end;

procedure TFVentAna.OKBtnClick(Sender: TObject);
Var i : integer ;
    TOBA : TOB ;
    Okok : boolean ;
    Section,RefA,Pref,NatureId : String ;
    Taux     : Double ;
    NumV,k   : integer ;
    FListe   : THGrid ;
    TotV     : Array[1..3] of Double ;
begin
  //C.B 16/06/2006
  SendMessage (FListe1.Handle, WM_KeyDown, VK_TAB, 0);
 
  if Action=taConsult then BEGIN Close ; Exit ; END ;
  TOBAna.ClearDetail ;
  if TOBL.NomTable='PAIEENCOURS' then
     BEGIN
     RefA:=LaReference ; NatureId:='PG' ;
     Pref:='PPU' ;
     END else
     BEGIN
     {$IFNDEF PAIEGRH}
     if TOBL.NomTable='PIECE' then RefA:=EncodeRefCPGescom(TOBL)
                              else RefA:=EncodeRefCPGescom(TOBL.Parent) ;
     {$ENDIF}
     if Ind=1 then Pref:='GS' else Pref:='GL' ;
     NatureId:='GC' ;
     END ;
  NumV:=0 ; FillChar(TotV, Sizeof(TotV),#0) ;
  for k:=1 to 3 do if TTabSheet(FindComponent('TS'+IntToStr(k))).TabVisible then
      BEGIN
      if k=1 then FListe:=FListe1 else if k=2 then FListe:=FListe2 else FListe:=FListe3 ;
      for i:=1 to FListe.RowCount-1 do
        BEGIN
        Section:=FListe.Cells[1,i] ;
        Taux:=Valeur(FListe.Cells[3,i]) ;
        if ((Section<>'') and (Taux<>0)) then
           BEGIN
           Inc(NumV) ;
           TOBA:=TOB.Create('VENTANA',TOBAna,-1) ;
           TOBA.PutValue('YVA_TABLEANA',Pref) ;
           TOBA.PutValue('YVA_IDENTIFIANT',RefA) ;
           TOBA.PutValue('YVA_AXE','A'+IntToStr(k)) ;
           TOBA.PutValue('YVA_NATUREID',NatureId) ;
           TOBA.PutValue('YVA_IDENTLIGNE',FormatFloat('000',0)) ;
           TOBA.PutValue('YVA_SECTION',Section) ;
           TOBA.PutValue('YVA_POURCENTAGE',Taux) ;
           TOBA.PutValue('YVA_NUMVENTIL',NumV) ;
           TotV[k]:=Arrondi(TotV[k]+Taux,2) ;
           END ;
        END ;
      END ;
  Okok:=True ;
  for k:=1 to 3 do
  begin
    if ((TotV[k]>100.0) or (TotV[k]<=0)) then
      if TTabSheet(FindComponent('TS'+IntToStr(k))).TabVisible then
      BEGIN
        Okok:=False ;
        Break ;
      END ;
  end;
  if Not Okok then
   { N° msg erreur : 10=axe1, 11=axe2, 12=axe3 }
    if MsgBox.Execute(9+k,'','')<>mrYes then
      Exit ;
    retOk := true;
  Close ;
end;

procedure TFVentAna.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
StCellCur:='' ; FClosing:=False ;
end;

{$IFNDEF EAGLSERVER}
{$IFNDEF EAGLCLIENT}
Function TFVentAna.ChercheSect( GS : THGrid ; C,L : integer ) : byte ;
Var
  St : String;
  CSect : TGSection;
  StType : String;
  Cache : THCpteEdit;

BEGIN

	ChercheSect:=0;
	St:=uppercase(GS.Cells[C,L]);
	Cache:=THCpteEdit.Create(Self);
	Cache.Parent := GS;//mcd 20/09/04 suite conseil EPZ pour fiche 10917 AGL et 11794 GC
	Cache.Visible:=False;
  Cache.Libelle:=TCache;
	Cache.Text:=St ;
	if GS=FListe1 then
  begin
    Cache.ZoomTable:=tzSection;
    if VH^.LiaisonY2ViaShare then StType:='tzCSection'  else StType:='tzSection' ;
	end
  else if GS=FListe2 then
  begin
    Cache.ZoomTable:=tzSection2;
    if VH^.LiaisonY2ViaShare then StType:='tzCSection2'  else StType:='tzSection2' ;
  end
  else
  begin
    Cache.ZoomTable:=tzSection3;
    if VH^.LiaisonY2ViaShare then StType:='tzCSection3'  else StType:='tzSection3' ;
  end;

	if GChercheCompte(Cache,Nil) then
  BEGIN
   	if St<>Cache.Text then
   	BEGIN
      GS.Cells[C,L]:=Cache.Text;
      GS.Cells[2,L]:=TCache.Caption;
      ChercheSect:=1;
      CSect:=TGSection.Create(GS.Cells[C,L],Cache.ZoomTable);
      if ((CSect<>Nil) and (CSect.Ferme)) then MsgBox.Execute(1,'','');
      CSect.Free;
      // afficher le libellé
		  GS.Cells[2,L]:= RechDom(StType,GS.Cells[C,L],False);
    END
    else
    BEGIN
      ChercheSect:=2;
   	END ;
  END ;
	Cache.Free;
END;

{$ELSE}
Function TFVentAna.ChercheSect( GS : THGrid ; C,L : integer ) : byte ;
Var St,sWhere : String ;
    NumA,OldCol,OldRow : integer ;
    bb   : boolean ;
    StType : String ;
BEGIN
  if VH^.LiaisonY2ViaShare then
  begin
    ChercheSect:=0 ;
    St:=uppercase(GS.Cells[C,L]) ;
    if GS=FListe1 then BEGIN NumA:=1 ; StType:='tzCSection' ; END else
    if GS=FListe2 then BEGIN NumA:=2 ; StType:='tzCSection2' ; END else
    BEGIN NumA:=3 ; StType:='tzCSection3' ; END ;
    sWhere:='CSP_SECTION LIKE "'+St+'%" AND CSP_AXE="A'+IntToStr(NumA)+'" AND CSP_FERME<>"X"' ;
    OldCol:=GS.Col ; OldRow:=GS.Row ; GS.Col:=C ; GS.Row:=L ;
    bb:=LookupList(GS,MsgBox.Mess[7],'CSECTION','CSP_SECTION','CSP_LIBELLE',sWhere,'CSP_SECTION',True,0) ;
    if bb then
    BEGIN
      ChercheSect:=1 ;
      GS.Cells[2,L]:=RechDom(StType,GS.Cells[C,L],False) ;
      GS.Col:=OldCol ; GS.Row:=OldRow ;
    END ;
  end else
  begin
    ChercheSect:=0 ;
    St:=uppercase(GS.Cells[C,L]) ;
    if GS=FListe1 then BEGIN NumA:=1 ; if VH^.LiaisonY2ViaShare then StType:='tzCSection'  else StType:='tzSection' ; END else
    if GS=FListe2 then BEGIN NumA:=2 ; if VH^.LiaisonY2ViaShare then StType:='tzCSection2'  else StType:='tzSection2' ; END else
    BEGIN NumA:=3 ; if VH^.LiaisonY2ViaShare then StType:='tzCSection3'  else StType:='tzSection3'; END ;
    if VH^.LiaisonY2ViaShare then
    begin
      sWhere:='CSP_SECTION LIKE "'+St+'%" AND CSP_AXE="A'+IntToStr(NumA)+'" AND CSP_FERME<>"X"' ;
    end else
    begin
      sWhere:='S_SECTION LIKE "'+St+'%" AND S_AXE="A'+IntToStr(NumA)+'" AND S_FERME<>"X"' ;
    end;
    OldCol:=GS.Col ; OldRow:=GS.Row ; GS.Col:=C ; GS.Row:=L ;
    if VH^.LiaisonY2ViaShare then
    begin
      bb:=LookupList(GS,MsgBox.Mess[7],'CSECTION','CSP_SECTION','CSP_LIBELLE',sWhere,'CSP_SECTION',True,0) ;
    end else
    begin
      bb:=LookupList(GS,MsgBox.Mess[7],'SECTION','S_SECTION','S_LIBELLE',sWhere,'S_SECTION',True,0) ;
    end;
    if bb then
    BEGIN
      ChercheSect:=1 ;
      GS.Cells[2,L]:=RechDom(StType,GS.Cells[C,L],False) ;
      GS.Col:=OldCol ; GS.Row:=OldRow ;
    END ;
  end;
END ;
{$ENDIF}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
procedure TFVentAna.FListe1DblClick(Sender: TObject);
Var F  : THGrid ;
    b  : byte ;
{$IFNDEF EAGLCLIENT}
    A  : TActionFiche ;
{$IFNDEF BTP}
    NumA : integer ;
{$ENDIF}
{$ENDIF EAGLCLIENT}
begin
F:=THGrid(Sender) ;
{$IFNDEF EAGLCLIENT}
{$IFNDEF BTP}
if F=Fliste1 then NumA:=1 else if F=Fliste2 then NumA:=2 else NumA:=3 ;
{$ENDIF}
{ En provenance de la GC, on a pas le droit de modifier }
if not FromGC then
begin
  if ((Action=taConsult) or (Not ExJaiLeDroitConcept(TConcept(ccSecModif),False))) then
    A:=taConsult
    else
    A:=taModif ;
end else
  A := taConsult;
{$ENDIF EAGLCLIENT}
if F.Col=1 then
   BEGIN
   b:=ChercheSect(F,F.Col,F.Row) ;
   if b=2 then
      BEGIN
{$IFDEF EAGLCLIENT}
// AFAIREEAGL
{$ELSE}
	{$IFNDEF BTP}
      FicheSection(Nil,'A'+IntToStr(NumA),F.Cells[F.Col,F.Row],A,0) ;
  {$ENDIF}
{$ENDIF}
      END else if ((Action<>taConsult) and (F.Cells[1,F.Row]<>'')) then InitTaux(F,F.Row) ;
   END ;
end;
{$ENDIF EAGLSERVER}

procedure TFVentAna.FListe1CellExit(Sender: TObject; var ACol,ARow: Longint; var Cancel: Boolean);
Var GS : THGrid ;
begin
if csDestroying in ComponentState then Exit ;
if Action=taConsult then Exit ;
GS:=THGrid(Sender) ;
if GS.Cells[ACol,ARow]=StCellCur then Exit ;
if (ACol=1) then GS.Cells[ACol,ARow]:=UpperCase(GS.Cells[ACol,ARow]) ;
if Not CodeSectOk then
   BEGIN
   if Not SectionExiste(Sender,ARow) then
      BEGIN
{$IFDEF EAGLCLIENT}
      Cancel:=True ;
{$ELSE}
{$ENDIF}
{$IFNDEF EAGLSERVER}
      if ChercheSect(GS,1,ARow)<=0 then
         BEGIN
         Cancel:=True ; Exit ;
         END else
         BEGIN
         Cancel:=False ;
         CodeSectOk:=True ;
         if GS.Cells[1,ARow]<>''then InitTaux(GS,Arow) ;
         END ;
{$ENDIF EAGLSERVER}
      END ;
   END ;
if ACol>2 then GS.Cells[ACol,ARow]:=StrFMontant(Valeur(GS.Cells[ACol,ARow]),15,4,'',TRUE) ;
Recalc(GS) ;
end;

Procedure TFVentAna.NumeroteLigne(Sender : TObject) ;
Var i : Integer ;
BEGIN
for i:=1 to THGrid(Sender).RowCount-1 do THGrid(Sender).Cells[0,i]:=IntToStr(i) ;
END ;

procedure TFVentAna.FListe1CellEnter(Sender: TObject; var ACol,ARow: Longint; var Cancel: Boolean);
Var G : THGrid ;
begin
if Action=taConsult then Exit ;
G:=THGrid(Sender) ;
if G.Col=2 then
   BEGIN
   if Acol=3 then G.Col:=1 ;
   if Acol=1 then G.Col:=3 ;
   END ;
if Not Cancel then StCellCur:=G.Cells[G.Col,G.Row] ;
end;

procedure TFVentAna.FTypeClick(Sender: TObject);
Var Q : TQuery ;
    Section,StDom : String ;
    ii      : integer ;
    FListe  : THGrid ;
BEGIN
  if Action=taConsult then Exit ;
  if FType.Value='' then Exit ;
  if VH^.LiaisonY2ViaShare then
  begin
    for ii:=1 to 3 do
    BEGIN
      if ii=1 then BEGIN FListe:=FListe1 ; StDom:='TZCSECTION' ; END else
      if ii=2 then BEGIN FListe:=FListe2 ; StDom:='TZCSECTION2' ; END else
      BEGIN FListe:=FListe3 ; StDom:='TZCSECTION3' ; END ;
      FListe.VidePile(True) ;
      Q:=OpenSQL('SELECT * FROM VENTIL WHERE V_NATURE="TY'+IntToStr(ii)+'" AND V_COMPTE="'+FType.Value+'" ORDER BY V_NUMEROVENTIL',True) ;
      While Not Q.EOF do
      BEGIN
        FListe.Cells[0,FListe.RowCount-1]:=IntToStr(Q.FindField('V_NUMEROVENTIL').AsInteger) ;
        Section:=Q.FindField('V_SECTION').AsString ;
        FListe.Cells[1,FListe.RowCount-1]:=Section ;
        FListe.Cells[2,FListe.RowCount-1]:=RechDom('TZCSECTION',Section,False) ;
        FListe.Cells[3,FListe.RowCount-1]:=StrFMontant(Q.FindField('V_TAUXMONTANT').AsFloat,15,4,'',TRUE) ;
        FListe.RowCount:=FListe.RowCount+1 ;
        Q.Next ;
      END ;
      Ferme(Q) ;
    END ;
  end else
  begin
    for ii:=1 to 3 do
    BEGIN
      if ii=1 then BEGIN FListe:=FListe1 ; StDom:='TZSECTION' ; END else
      if ii=2 then BEGIN FListe:=FListe2 ; StDom:='TZSECTION2' ; END else
      BEGIN FListe:=FListe3 ; StDom:='TZSECTION3' ; END ;
      FListe.VidePile(True) ;
      Q:=OpenSQL('SELECT * FROM VENTIL WHERE V_NATURE="TY'+IntToStr(ii)+'" AND V_COMPTE="'+FType.Value+'" ORDER BY V_NUMEROVENTIL',True) ;
      While Not Q.EOF do
      BEGIN
        FListe.Cells[0,FListe.RowCount-1]:=IntToStr(Q.FindField('V_NUMEROVENTIL').AsInteger) ;
        Section:=Q.FindField('V_SECTION').AsString ;
        FListe.Cells[1,FListe.RowCount-1]:=Section ;
        FListe.Cells[2,FListe.RowCount-1]:=RechDom('TZSECTION',Section,False) ;
        FListe.Cells[3,FListe.RowCount-1]:=StrFMontant(Q.FindField('V_TAUXMONTANT').AsFloat,15,4,'',TRUE) ;
        FListe.RowCount:=FListe.RowCount+1 ;
        Q.Next ;
      END ;
      Ferme(Q) ;
    END ;
  end;
end;

procedure TFVentAna.BFermerClick(Sender: TObject);
begin
Close ;
if FClosing and IsInside(Self) then THPanel(parent).CloseInside ;
end;

procedure TFVentAna.FListe1RowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin CodeSectOk:=False ; end;

Function TFVentAna.SectionExiste(Sender :  TObject ; ARow : Integer) : Boolean ;
Var StC : String ;
    Q : TQuery ;
    NumA : integer ;
BEGIN
CodeSectOk:=False ; Result:=False ;
StC:=THGrid(Sender).Cells[1,ARow] ; if Stc='' then Exit ;
if Sender=FListe1 then NumA:=1 else if Sender=FListe2 then NumA:=2 else NumA:=3 ;
Q:=OpenSql('Select S_SECTION From SECTION Where S_SECTION="'+StC+'" And S_AXE="A'+IntToStr(NumA)+'"',True) ;
Result:=Not Q.Eof ;
Ferme(Q) ; CodeSectOk:=Result ;
END ;

procedure TFVentAna.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FListe1.VidePile(True) ;
FListe2.VidePile(True) ;
FClosing:=True ;
end;

procedure TFVentAna.HelpBtnClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

procedure TFVentAna.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
Case Key of
   VK_RETURN : if Screen.ActiveControl is THGrid then Key:=VK_TAB ;
   VK_F10    : BEGIN Key:=0 ; OkBtnClick(Nil) ; END ;
   VK_F6     : BEGIN Key:=0 ; BSoldeClick(Nil) ; END ;
   END ;
end;

procedure TFVentAna.PagesChange(Sender: TObject);
begin
if Pages.Activepage=TS1 then Recalc(FListe1) else if Pages.Activepage=TS2 then Recalc(FListe2) else Recalc(FListe3) ;
end;

procedure TFVentAna.BSoldeClick(Sender: TObject);
Var FL : THGrid ;
    XX,Delta : Double ;
    i  : integer ;
begin
if Pages.ActivePage=TS1 then FL:=FListe1 else if Pages.ActivePage=TS2 then FL:=FListe2 else FL:=FListe3 ;
if FL.Cells[1,FL.Row]='' then Exit ;
XX:=0 ;
For i:=1 to FL.RowCount-1 do XX:=XX+Valeur(FL.Cells[3,i]) ;
Delta:=Arrondi(100.0-XX,6) ; if Delta=0 then Exit ;
FL.Cells[3,FL.Row]:=StrFMontant(Valeur(FL.Cells[3,FL.Row])+Delta,15,4,'',TRUE) ;
Recalc(FL) ;
end;

//C.B 16/06/2006
procedure TFVentAna.FListe1ColExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  CodeSectOk := false;
end;

procedure TFVentAna.FListe1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
	case key of
    VK_DELETE: if (Shift = [ssCtrl]) then
    begin
      Key := 0;
      DeleteLigne;
    end;
  end;
end;

procedure TFVentAna.DeleteLigne;
begin
	if FListe1.row > 0 then FListe1.DeleteRow(FListe1.row);
end;

end.
