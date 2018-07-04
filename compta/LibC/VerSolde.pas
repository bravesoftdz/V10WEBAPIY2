{***********UNITE*************************************************
Auteur  ...... : BPY
Créé le ...... : 24/09/2004
Modifié le ... : 24/09/2004
Description .. : - BPY le 24/09/2004 - Passage eAGL
Mots clefs ... : 
*****************************************************************}
unit VerSolde;

interface

uses
    Windows,
    Messages,
    SysUtils,
    Classes,
    Graphics,
    Controls,
    Forms,
    Dialogs,
    Hctrls,
    StdCtrls,
    Mask,
    ExtCtrls,
    ComCtrls,
    Grids,
    Buttons,
    Ent1,
    HEnt1,
    HStatus,
    cpteutil,
{$IFDEF EAGLCLIENT}
    UTOB,
{$ELSE}
    HDB,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
    CalcOle,
    HSysMenu,
    Menus,
    UObjFiltres,
    HTB97,
    {$IFDEF MODENT1}
    CPTypeCons,
    {$ELSE}
    tCalcCum,
    {$ENDIF MODENT1}
    HmsgBox
    ;

procedure ControleSolde;

type
    TFVerSolde = class(TForm)
        PanelBouton : TPanel;
    BReduire: THBitBtn;
    BAgrandir: THBitBtn;
    BRechercher: THBitBtn;
        Panel1 : TPanel;
    BImprimer: THBitBtn;
    BFermer: THBitBtn;
    BAide: THBitBtn;
        Pages : TPageControl;
        PCritere : TTabSheet;
        Bevel1 : TBevel;
        FCptes : TEdit;
        TFCptes : THLabel;
        PFiltres : TPanel;
    BCherche: THBitBtn;
        FFiltres : THValComboBox;
        Formateur : THNumEdit;
        FListe : THGrid;
        FindDialog : TFindDialog;
        HMTrad : THSystemMenu;
        POPF : TPopupMenu;
        BCreerFiltre : TMenuItem;
        BSaveFiltre : TMenuItem;
        BDelFiltre : TMenuItem;
        BRenFiltre : TMenuItem;
        BNouvRech : TMenuItem;
        BFiltre : TToolbarButton97;
        procedure BAgrandirClick(Sender : TObject);
        procedure BChercheClick(Sender : TObject);
        procedure FormShow(Sender : TObject);
        procedure BCreerFiltreClick(Sender : TObject);
        procedure BSaveFiltreClick(Sender : TObject);
        procedure BDelFiltreClick(Sender : TObject);
        procedure BRenFiltreClick(Sender : TObject);
        procedure BNouvRechClick(Sender : TObject);
        procedure FindDialogFind(Sender : TObject);
        procedure BRechercherClick(Sender : TObject);
        procedure BImprimerClick(Sender : TObject);
        procedure BFermerClick(Sender : TObject);
        procedure FormCreate(Sender : TObject);
        procedure BReduireClick(Sender : TObject);
        procedure FFiltresChange(Sender : TObject);
        procedure BAideClick(Sender : TObject);
        procedure POPFPopup(Sender : TObject);
    private
    { Déclarations privées }
        GCAlc : TGCalculCum;
        FindFirst : Boolean;
        Nomfiltre : string;
        WMinX, WMinY : Integer;
        ObjetFiltre  : TObjFiltre;
        procedure LanceTraitement;
        function SaisieOk : Boolean;
        procedure AffecteGrille(Q : TQuery;Tot1, Tot2 : TabTot;I : Integer);
        procedure WMGetMinMaxInfo(var MSG : Tmessage); message WM_GetMinMaxInfo;
    public
    { Déclarations publiques }
    end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
{$IFDEF EAGLCLIENT}
    UtileAGL,
{$ELSE}
    PrintDBG,
{$ENDIF}
    UtilPgi
    ;

procedure ControleSolde;
var
    CSold : TFVerSolde;
begin
    if not _BlocageMonoPoste(True) then Exit;
    CSold := TFVerSolde.Create(Application);
    try
        CSold.ShowModal;
    finally
        CSold.Free;
        _DeblocageMonoPoste(True);
    end;
    Screen.Cursor := SyncrDefault;
end;

procedure TFVerSolde.BAgrandirClick(Sender : TObject);
begin
    ChangeListeCrit(Self, True);
end;

procedure TFVerSolde.BChercheClick(Sender : TObject);
begin
    FListe.RowCount := 2;
    FListe.Cells[0, 1] := '';
    FListe.Cells[1, 1] := '';
    FListe.Cells[2, 1] := '';
    EnableControls(Self, False);
    if SaisieOk then
    begin
        LanceTraitement;
        GCalc.Free;
    end;
    EnableControls(Self, True);
end;

procedure TFVerSolde.WMGetMinMaxInfo(var MSG : Tmessage);
begin
    with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do
    begin
        X := WMinX;
        Y := WMinY;
    end;
end;

function TFVerSolde.SaisieOk : Boolean;
begin
    Result := (FCptes.Text <> '');
    if Result then
    begin
          //TGCalculCum.Create(FMultiple : TTypeCalc ; Ffb1,Ffb2 : TFichierBase ; FSetTyp : SetttTypePiece ; FFiltreDev,FFiltreEtab,FFiltreExo,FDeviseEnPivot,FEnEuro : Boolean ; FDec,FDecE : Integer) ;
        Gcalc := TGCalculCum.create(Un, fbGene, fbGene, [tpReel], False, False, True, False, False, V_PGI.OkDecV, V_PGI.OkDecE);
        GCalc.initCalcul('', '', '', '', '', '', VH^.Encours.Code,
            VH^.EnCours.Deb, VH^.EnCours.Deb, TRUE);
   (**)
    end;
end;

procedure TFVerSolde.LanceTraitement;
var
    TotCpt1, TotCpt2 : TabTot;
    StComptes, S, St : string;
    Rempli, NbMvt : Integer;
    Q : Tquery;
begin
    Fillchar(TotCpt1, SizeOf(TotCpt1), #0);
    Fillchar(TotCpt2, SizeOf(TotCpt2), #0);
    InitMove(100, '');
    NbMvt := 0;
    Rempli := 0;
    StComptes := AnalyseCompte(FCptes.Text, fbGene, False, False);
    St := ' Select distinct(G_GENERAL), G_LIBELLE from GENERAUX where ' +
        ' ' + StComptes + ' order by g_general';
    Q := OpenSql(St, True);
    while not Q.Eof do
    begin
        MoveCur(FALSE);
        S := Q.Fields[0].AsString;
        GCAlc.ReInitCalcul(Q.Fields[0].AsString, '', VH^.EnCours.Deb, VH^.EnCours.Fin);
        GCalc.Calcul;
        TotCpt1 := GCalc.ExecCalc.TotCpt;

(*
      GCAlc.ReInitCalcul(Q.Fields[0].AsString, '',VH^.EnCours.Deb,VH^.EnCours.Fin) ;
      GCalc.Calcul ; TotCpt2:=GCalc.ExecCalc.TotCpt ;
*)
        CumulVersSolde(TotCpt1[0]);
        inc(Rempli);
        AffecteGrille(Q, TotCpt1, TotCpt2, Rempli);
        if NbMvt >= 100 then
        begin
            NbMvt := 0;
            FiniMove;
            InitMove(100, '');
        end;
        Q.Next;
    end;
    Ferme(Q);
    if FListe.RowCount > 2 then FListe.RowCount := FListe.RowCount - 1;
    FiniMove;
end;

procedure TFVerSolde.AffecteGrille(Q : TQuery;Tot1, Tot2 : TabTot;I : Integer);
var
    M1 : TabTot;
begin
    Application.ProcessMessages;
    FListe.Cells[0, i] := Q.FindField('G_GENERAL').AsString;
    M1[1].TotDebit := Arrondi(ToT1[0].TotDebit + ToT1[1].TotDebit + Tot2[1].TotDebit, V_PGI.OkDecV);
    M1[1].TotCredit := Arrondi(ToT1[0].TotCredit + ToT1[1].TotCredit + Tot2[1].TotCredit, V_PGI.OkDecV);
    FListe.Cells[1, i] := Q.FindField('G_LIBELLE').AsString;
    FListe.Cells[2, i] := PrintSolde(M1[1].TotDebit, M1[1].TotCredit, V_PGI.OkDecV, V_PGI.SymbolePivot, True);
    FListe.RowCount := FListe.RowCount + 1;
end;

procedure TFVerSolde.FormShow(Sender : TObject);
var
TCF        : TControlFiltre;
begin
    PopUpMenu := ADDMenuPop(PopUpMenu, '', '');
    ChangeMask(Formateur, V_PGI.OkDecV, '');
    Nomfiltre := 'CONTOLSOLDE';
//    ChargeFiltre(Nomfiltre, FFiltres, Pages);
    // fiche 18816
    TCF.Filtre := BFiltre;
    TCF.Filtres := FFiltres;
    TCF.PageCtrl := Pages;
    ObjetFiltre := TObjFiltre.create(TCF, 'VERSOLDE');
    ObjetFiltre.Charger;
end;

procedure TFVerSolde.BCreerFiltreClick(Sender : TObject);
begin
//    NewFiltre(Nomfiltre, FFiltres, Pages);
end;

procedure TFVerSolde.BSaveFiltreClick(Sender : TObject);
begin
//    SaveFiltre(Nomfiltre, FFiltres, Pages);
end;

procedure TFVerSolde.BDelFiltreClick(Sender : TObject);
begin
//    DeleteFiltre(Nomfiltre, FFiltres);
end;

procedure TFVerSolde.BRenFiltreClick(Sender : TObject);
begin
//    RenameFiltre(Nomfiltre, FFiltres);
end;

procedure TFVerSolde.BNouvRechClick(Sender : TObject);
begin
//    VideFiltre(FFiltres, Pages);
end;

procedure TFVerSolde.FindDialogFind(Sender : TObject);
begin
    Rechercher(FListe, FindDialog, FindFirst);
end;

procedure TFVerSolde.BRechercherClick(Sender : TObject);
begin
    FindFirst := True;
    FindDialog.Execute;
end;

procedure TFVerSolde.BImprimerClick(Sender : TObject);
begin
{$IFDEF EAGLCLIENT}
{$ELSE}
    PrintDBGrid(FListe, Pages, Caption, '');
{$ENDIF}
end;

procedure TFVerSolde.BFermerClick(Sender : TObject);
begin
    if Assigned(ObjetFiltre) then FreeAndNil(ObjetFiltre);
    Close;
end;

procedure TFVerSolde.FormCreate(Sender : TObject);
begin
    WMinX := Width;
    WMinY := 320;
end;

(*
M1[2]:=M1[1] ;
CumulVersSolde(M1[2]) ;
if M1[2].TotDebit=0 then
   BEGIN
   FListe.Cells[1,i]:='' ;
   FListe.Cells[2,i]:=AfficheMontant(Formateur.Masks.PositiveMask,V_PGI.SymbolePivot,M1[2].TotCredit,True) ;
   END else
   BEGIN
   FListe.Cells[1,i]:=AfficheMontant(Formateur.Masks.PositiveMask,V_PGI.SymbolePivot,M1[2].TotDebit,True ) ;
   FListe.Cells[2,i]:='' ;
   END ;
*)

procedure TFVerSolde.BReduireClick(Sender : TObject);
begin
    ChangeListeCrit(Self, False);
end;

procedure TFVerSolde.FFiltresChange(Sender : TObject);
begin
  //  LoadFiltre(NomFiltre, FFiltres, Pages);
end;

procedure TFVerSolde.BAideClick(Sender : TObject);
begin
    CallHelpTopic(Self);
end;

procedure TFVerSolde.POPFPopup(Sender : TObject);
begin
 //   UpdatePopFiltre(BSaveFiltre, BDelFiltre, BRenFiltre, FFiltres);

end;

end.
