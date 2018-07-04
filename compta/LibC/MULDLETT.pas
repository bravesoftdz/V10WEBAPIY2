unit MULDLETT;

//================================================================================
// Interface
//================================================================================
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
{$IFNDEF EAGLCLIENT}
    DB,
    {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
    Grids,
    StdCtrls,
    DBCtrls,
    Mask,
    Hctrls,
    Hcompte,
    SaisUtil,
    ComCtrls,
    Buttons,
    ExtCtrls,
    Hqry,
    Ent1,
    LettUtil,
    Lettrage,
    hmsgbox,
    HEnt1,
    Menus,
//    Filtre,
    HDB,
    HSysMenu,
    HTB97,
    HPanel,
    UiUtil,
    UTOB,
    LookUp,
    UObjFiltres {JP 04/10/04 : Gestion des filtres}
    ;

//==================================================
// Externe
//==================================================
procedure LanceDeLettrage;
procedure LanceDeLettrageMP(Qui:tProfilTraitement);

//==================================================
// Definition de class
//==================================================
type
  TFMulDLettre = class(TForm)

        Pages : TPageControl;
        Princ : TTabSheet;
        PLibres : TTabSheet;

        GL : THGrid;

        FindLettre : TFindDialog;
        HMessMulDL : THMsgBox;
        HMTrad : THSystemMenu;

        FFiltres : THValComboBox;

        Cache : THCpteEdit;
    BFeuVert: THBitBtn;
        CFacNotPay : TCheckBox;

        BChercher : TToolbarButton97;

        E_GENERAL : THCritMaskEdit;
        E_GENERAL_ : THCritMaskEdit;
        E_AUXILIAIRE : THCritMaskEdit;
        E_AUXILIAIRE_ : THCritMaskEdit;
        LETTRAGE : THCritMaskEdit;
        LETTRAGE_ : THCritMaskEdit;
        E_REFLETTRAGE : THCritMaskEdit;
       iCritGlyph: TImage;
       iCritGlyphModified: TImage;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    BFiltre: TToolbarButton97;

        // evenement par default
        procedure FormCreate(Sender:TObject);
        procedure FormShow(Sender:TObject);
        procedure FormClose(Sender:TObject;var Action:TCloseAction);
        // keyDown
        procedure FormKeyDown(Sender:TObject;var Key:Word;Shift:TShiftState);
        // click de la bar
        procedure BAgrandirClick(Sender:TObject);
        procedure BReduireClick(Sender:TObject);
        procedure BRechercheClick(Sender:TObject);
        procedure BImprimerClick(Sender:TObject);
        procedure BExecClick(Sender:TObject);
        procedure BValiderClick(Sender:TObject);
        procedure BFermeClick(Sender:TObject);
        procedure BAideClick(Sender:TObject);
        // Click de la bar de recherche
        procedure BChercherClick(Sender:TObject);
        procedure BFeuVertClick(Sender:TObject);
        // dbl click
        procedure GLDblClick(Sender:TObject);
        // elipsis Click
        procedure E_GENERALElipsisClick(Sender : TObject);
        procedure E_AUXILIAIREElipsisClick(Sender : TObject);
        // change
        // exit
        procedure E_GENERALExit(Sender:TObject);
        procedure E_AUXILIAIREExit(Sender : TObject);
        // autre evenement
        procedure FindLettreFind(Sender:TObject);
    private
        QL : TQuery;
        WMinX, WMinY:Integer;
        LesCriteres:string;
        Qui:tProfilTraitement;
        Loading, FCritModified: boolean;

        // message
        procedure WMGetMinMaxInfo(var MSG:Tmessage); message WM_GetMinMaxInfo;
        // creation de la requete
        procedure RechercheEcritures;
        function WhereCrit:string;
        function CritTP:string;
        // control
//        procedure ControleFeuVert;

        procedure AttribCollectif;
        procedure NextAuxi(Suiv:boolean);
        procedure NextGene(Suiv:boolean);

        // traitement
        procedure DeLettrage;
        procedure SetCritModified(Value: Boolean);
        procedure CritereChange(Sender: TObject);
    protected
        ObjFiltre : TObjFiltre;

        procedure Loaded; override;
    public
        FindFirst:boolean;
        property CritModified: Boolean read FCritModified write SetCritModified;
    end;

//================================================================================
// Implementation
//================================================================================
implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPTypeCons,
  {$ENDIF MODENT1}
    ULibWindows,
    ParamDBG,
{$IFNDEF EAGLCLIENT}
    PrintDBG,
{$ENDIF}
    UtilPgi
    ;

//==================================================
// fonctions hors class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure LanceDeLettrage;
var
    X:TFMulDLettre;
    PP:THPanel;
begin
    if (_Blocage(['nrCloture'],false,'nrAucun')) then exit;

    X := TFMulDLettre.Create(Application);
    PP := FindInsidePanel;

    if (PP = nil) then
    begin
        try
            X.Qui := prAucun;
            X.ShowModal;
        finally
            X.Free;
        end;

        Screen.Cursor := SyncrDefault;
    end
    else
    begin
        X.Qui := prAucun;
        InitInside(X, PP);
        X.Show;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure LanceDeLettrageMP(Qui:tProfilTraitement);
var
    X:TFMulDLettre;
    PP:THPanel;
begin
    if (_Blocage(['nrCloture'],false,'nrAucun')) then exit;

    X := TFMulDLettre.Create(Application);
    PP := FindInsidePanel;

    if (PP = nil) then
    begin
        try
            X.Qui := Qui;
            X.ShowModal;
        finally
            X.Free;
        end;

        Screen.Cursor := SyncrDefault;
    end
    else
    begin
        X.Qui := Qui;
        InitInside(X, PP);
        X.Show;
    end;
end;

//==================================================
// Evenements de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.FormCreate(Sender:TObject);
var
  Composants : TControlFiltre;
begin
    Loading := TRUE;
    PopUpMenu := ADDMenuPop(PopUpMenu,'','');
    WMinX := Width;
    WMinY := 229;
  {JP 04/10/04 : Gestion des filtres}
  Composants.PopupF   := POPF;
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := Pages;
  ObjFiltre := TObjFiltre.Create(Composants, 'DELETTRAGE');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 06/02/2002
Modifié le ... :   /  /
Description .. : on autorise le multiselection sur la grille
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.FormShow(Sender:TObject);
begin
    // proprieté de la grille
    GL.ListeParam := 'MULDLETEUR';
    // nb D
    GL.ColColors[4] := clBlue;
    // debit
    GL.ColColors[5] := clBlue;
    // nb c
    GL.ColColors[6] := clRed;
    // credit
    GL.ColColors[7] := clRed;
    // sold
    GL.ColColors[8] := clGreen;

    LesCriteres := '';
    InitTablesLibresTiers(PLibres);

{$IFNDEF CCS3}
    if (not VH^.OuiTP) then
{$ENDIF}
        CFacNotPay.Visible := false;

    HMTrad.ResizeGridColumns(GL);
    E_AUXILIAIRE.SetFocus;
    Loading := False;

    Pages.ActivePageIndex := 0;

  {JP 04/10/04 : Gestion des filtres}
  ObjFiltre.Charger;
  if (FFiltres.Text <> '') then
    BChercherClick(BChercher);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.FormClose(Sender:TObject;var Action:TCloseAction);
begin
  {JP 04/10/04 : Gestion des filtres}
  if Assigned(ObjFiltre) then ObjFiltre.Free;
    Ferme(QL);
{$IFNDEF EAGLCLIENT}
    if (Parent is THPanel) then Action := caFree;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.FormKeyDown(Sender:TObject;var Key:Word;Shift:TShiftState);
var
    S:Boolean;
begin
    S := (Shift = [ssShift]);

    case Key of
        VK_F4:
            begin
                Key := 0;
                if (S) then NextGene(true)
                else NextAuxi(true);
            end;
        VK_F3:
            begin
                Key := 0;
                if (S) then NextGene(false)
                else NextAuxi(false);
            end;
        VK_F9:
            begin
                Key := 0;
                BChercherClick(nil);
            end;
        VK_F10:
            begin
                Key := 0;
                BValiderClick(nil);
            end;
        Ord('A'):
            begin
                if (Shift = [ssCtrl]) then
                begin
                    Key := 0;
                    GL.AllSelected := (not GL.AllSelected);
                end;
            end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.BAgrandirClick(Sender:TObject);
begin
    ChangeListeCrit(Self, True);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.BReduireClick(Sender:TObject);
begin
    ChangeListeCrit(Self, False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.BRechercheClick(Sender:TObject);
begin
    FindFirst := True;
    FindLettre.Execute;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.BImprimerClick(Sender:TObject);
begin
    if CritModified then
    begin
        PgiInfo('Vous n''avez pas appliqué les critères !');
        exit;
    end;
    {$IFDEF EAGLCLIENT}
    {$ELSE}
    PrintDBGrid(GL,nil,Caption,'');
    {$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 04/04/2002
Modifié le ... :   /  /
Description .. : Procedure qui lance le filtre, selectionne tous les
Suite ........ : enregistrements et lance le delettrage
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.BExecClick(Sender:TObject);
begin
    BChercher.Click;
    GL.AllSelected := not (GL.AllSelected);
    BValiderClick(nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 06/02/2002
Modifié le ... : 28/10/2002
Description .. : Si aucune ligne n'est selectionné on affiche la fenetre de
Suite ........ : delecttrage sinon on supprime le lettrage
Suite ........ : - 28/10/2002 - changement du message de confirmation
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.BValiderClick(Sender:TObject);
begin
  //LG* 05/01/2002
  if (PGIAsk('Confirmez-vous le traitement ?','Confirmation') = mrNo) then exit;

  if ((GL.nbSelected = 0) and (not GL.AllSelected)) then
    GLDblClick(nil)
  else
  begin
    Transactions(DeLettrage,1);
    BChercherClick(nil);
    GL.ClearSelected;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.BFermeClick(Sender:TObject);
begin
    if (IsInside(Self)) then CloseInsidePanel(self);
    Close;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.BAideClick(Sender:TObject);
begin
    CallHelpTopic(Self);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.BChercherClick(Sender:TObject);
{$IFNDEF EAGLCLIENT}
var
    tmpTOB : TOB;
{$ENDIF}
begin
    RechercheEcritures;
{$IFDEF EAGLCLIENT}
    QL.PutGridDetail(GL,false,false,'',true);
{$ELSE}
    tmpTOB := TOB.Create('',nil,-1);
    tmpTOB.LoadDetailDB('','','',QL,false,false);
    tmpTOB.PutGridDetail(GL,false,false,'',true);

    FreeAndNil(tmpTOB);
{$ENDIF}
    HMTrad.ResizeGridColumns(GL);
    CritModified := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.BFeuVertClick(Sender:TObject);
var
    Q:TQuery;
    Existe, Coll:Boolean;
begin
    Screen.Cursor := crHourglass;
    Coll := false;

    // controle auxi
    Q := OpenSQL('SELECT T_AUXILIAIRE, T_CONFIDENTIEL FROM TIERS WHERE T_AUXILIAIRE="' + E_AUXILIAIRE.Text + '"',true);
    if (not Q.EOF) then
    begin
        if (EstConfidentiel(Q.Fields[1].AsString)) then Existe := false
        else Existe := true;
    end
    else Existe := false;
    Ferme(Q);

    if (not Existe) then
    begin
        Screen.Cursor := SyncrDefault;
        HMessMulDL.Execute(2, '', '');
        E_AUXILIAIRE.SetFocus;
        Exit;
    end;

    // controle gene
    Q := OpenSQL('SELECT G_COLLECTIF, G_CONFIDENTIEL FROM GENERAUX WHERE G_GENERAL="' + E_GENERAL.Text + '"',true);
    if (not Q.EOF) then
    begin
        Coll := (Q.Fields[0].AsString = 'X');
        if (EstConfidentiel(Q.Fields[1].AsString)) then Existe := false
        else Existe := true;
    end
    else Existe := false;
    Ferme(Q);

    if (not Existe) then
    begin
        Screen.Cursor := SyncrDefault;
        HMessMulDL.Execute(1, '', '');
        E_GENERAL.SetFocus;
        Exit;
    end;
    if (not Coll) then
    begin
        Screen.Cursor := SyncrDefault;
        HMessMulDL.Execute(3, '', '');
        E_GENERAL.SetFocus;
        Exit;
    end;

    E_GENERAL_.Text := E_GENERAL.Text;
    E_AUXILIAIRE_.Text := E_AUXILIAIRE.Text;

{constitution lettrage}
    RechercheEcritures;
    if (QL.EOF) then
    begin
        Screen.Cursor := SyncrDefault;
        HMessMulDL.Execute(4, '', '');
        E_AUXILIAIRE.SetFocus;
    end
    else GLDblClick(nil);
    if (E_AUXILIAIRE_.Text <> '') or (E_GENERAL_.Text <> '') then
    begin
        E_AUXILIAIRE_.Text := '';
        E_GENERAL_.Text := '';

        Application.ProcessMessages;
        BChercherClick(nil);
    end;
    Screen.Cursor := SyncrDefault;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.GLDblClick(Sender:TObject);
var R:RLETTR;

begin
    // GCO - 20/07/2005 - FQ 15399
    if (QL = nil) OR ((QL.EOF) and (QL.BOF)) then Exit;

    FillChar(R, Sizeof(R), #0);
    R.General := GL.Cells[1,GL.Row];
    if (EstSQLConfidentiel('GENERAUX',R.General)) then
    begin
        HMessMulDL.Execute(1, '', '');
        Exit;
    end;
    R.Auxiliaire := GL.Cells[2,GL.Row];
    R.CritMvt := LesCriteres;
    R.Appel := tlMenu;
    R.CodeLettre := GL.Cells[9,GL.Row];
    R.DeviseMvt := GL.Cells[10,GL.Row];
    R.LettrageDevise := (GL.Cells[11,GL.Row] = 'X');
{$IFDEF CCMP}
    if (Qui <> prAucun) then LettrageManuelMP(R,false,taModif,Qui)
    else
{$ENDIF}
        LettrageManuel(R, False, taModif);
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 01/12/2004
Modifié le ... :   /  /    
Description .. : - LG - 01/12/2004 - FB 15017 - gestion des confidentiels
Mots clefs ... : 
*****************************************************************}
procedure TFMulDLettre.E_GENERALElipsisClick(Sender : TObject);
begin
  LookUpList(TControl(Sender),'Recherche d''un compte général','GENERAUX','G_GENERAL','G_LIBELLE','( G_LETTRABLE="X" OR G_COLLECTIF="X") AND ' + CGenereSQLConfidentiel('G') ,'G_GENERAL',true,-1);
//  ControleFeuVert;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 01/12/2004
Modifié le ... :   /  /    
Description .. : - LG - 01/12/2004 - FB 15017 - gestion des confidentiels
Mots clefs ... : 
*****************************************************************}
procedure TFMulDLettre.E_AUXILIAIREElipsisClick(Sender : TObject);
begin
  if (LookUpList(TControl(Sender),'Recherche d''un compte auxiliaire','TIERS','T_AUXILIAIRE','T_LIBELLE',CGenereSQLConfidentiel('T'),'T_AUXILIAIRE',true,-1)) then
    AttribCollectif;
//  ControleFeuVert;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 27/06/2006
Modifié le ... :   /  /
Description .. : - LG - FB 15017 - gestion des comptes confidentiels
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.E_GENERALExit(Sender:TObject);
begin
 if EstSQLConfidentiel('GENERAUX', THCritMaskEdit(Sender).Text) then
  begin
    HMessMulDL.Execute(1, '', '');
    THCritMaskEdit(Sender).Text := '' ;
    if THCritMaskEdit(Sender).Canfocus then
     THCritMaskEdit(Sender).SetFocus ;
    Exit;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 27/06/2006
Modifié le ... : 06/09/2007
Description .. : - LG - FB 15017 - gestion des comptes confidentiels
Mots clefs ... : 
*****************************************************************}
procedure TFMulDLettre.E_AUXILIAIREExit(Sender : TObject);
begin

 { FQ 20381 BVE 28.05.07 }
  if Trim(THEdit(Sender).Text) = '' then Exit;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if TestJoker(THEdit(Sender).Text) then Exit;

  // Complétion auto du numéro de compte si possible
  if ExisteSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE = "' + BourreEtLess(THEdit(Sender).Text, fbAux) + '"') then
  begin
     if Length(THEdit(Sender).Text) < VH^.Cpta[fbAux].Lg then
        THEdit(Sender).Text := BourreEtLess( THEdit(Sender).Text, fbAux);
  end
  else
     THEdit(Sender).ElipsisClick(nil);
  { END FQ 20381 }


 if EstSQLConfidentiel('TIERS', THCritMaskEdit(Sender).Text) then
  begin
    HMessMulDL.Execute(1, '', '');
    THCritMaskEdit(Sender).Text := '' ;
    if THCritMaskEdit(Sender).Canfocus then
     THCritMaskEdit(Sender).SetFocus ;
    Exit;
  end;
    AttribCollectif;
//    ControleFeuVert;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.FindLettreFind(Sender:TObject);
begin
    Rechercher(GL, FindLettre, FindFirst);
end;

//==================================================
// Gestion de message
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.WMGetMinMaxInfo(var MSG:Tmessage);
begin
    with (PMinMaxInfo(MSG.lparam)^.ptMinTrackSize) do
    begin
        X := WMinX;
        Y := WMinY;
    end;
end;

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.RechercheEcritures;
var
    Query : String;
begin
    Screen.Cursor := crHourglass;

    Ferme(QL);

    E_REFLETTRAGE.Text := '';
    LesCriteres := WhereCrit;

    Query := SelectQDL(2) + ' WHERE ' + LWhereBase(false,true,false,false) + LWhereComptes(Self) + LesCriteres + CritTP;
    if (MonProfilOk(Qui)) then Query := Query + ' AND ' + WhereProfilUser(QL,Qui);
    Query := Query + 'GROUP BY E_GENERAL, E_AUXILIAIRE, E_LETTRAGE';

    QL := OpenSQL(Query,true);

    Screen.Cursor := SyncrDefault;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TFMulDLettre.WhereCrit:string;
var
    StC:string;
    CC  : THCritMaskEdit;
    i:integer;
begin
    StC := Fourchette('UPPER(E_LETTRAGE)', LETTRAGE.Text, LETTRAGE_.Text, ftString);
    if (STC <> '') then StC := ' AND ' + StC;

    if (Plibres.TabVisible) then
    begin
        for i := 0 to 9 do
        begin
            CC := THCritMaskEdit(FindComponent('T_TABLE'+IntToStr(i)));
            if (CC = nil) then Continue;
            if (CC.Text <> '') then StC := StC + ' AND T_TABLE' + IntToStr(i) + '="' + CC.Text + '"';
        end;
    end;

    if (E_REFLETTRAGE.Text <> '') then StC := StC + ' AND E_REFLETTRAGE LIKE "' + E_REFLETTRAGE.Text + '%" ';

    result := StC;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TFMulDLettre.CritTP:string;
begin
    Result := '';

    if (not VH^.OuiTP) then exit;

    case CFacNotPay.State of
        cbChecked:Result := ' AND (E_TIERSPAYEUR="" OR E_TIERSPAYEUR IS NULL)';
        cbUnChecked:Result := ' AND E_TIERSPAYEUR<>"" ';
        cbGrayed:Result := '';
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
(*
procedure TFMulDLettre.ControleFeuVert;
var
    Okok:boolean;
begin
    Okok := True;

    if ((E_GENERAL_.Text <> '') or (E_AUXILIAIRE_.Text <> '')) then Okok := False;
    if ((E_GENERAL.Text = '') or (E_AUXILIAIRE.Text = '')) then Okok := False;

    // BPY le 17/12/2003 => suppression du btn "Feu vert" en delettrage
//    if (BFeuVert.Visible <> Okok) then BFeuVert.Visible := Okok;
    // fin BPY
end;
*)

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.AttribCollectif;
var
    Q:TQuery;
begin
    if ((E_AUXILIAIRE_.Text <> '') or (E_GENERAL.Text <> '') or (E_GENERAL_.Text <> '')) then Exit;

    Q := OpenSQL('Select T_COLLECTIF from Tiers Where T_AUXILIAIRE="' + E_AUXILIAIRE.Text + '"', True);

    if (not Q.EOF) then E_GENERAL.Text := Q.Fields[0].AsString;

    Ferme(Q);
//    ControleFeuVert;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.NextAuxi(Suiv:boolean);
var
    Q:TQuery;
    StConf:string;
begin
    if ((E_GENERAL_.Text <> '') or (E_AUXILIAIRE_.Text <> '') or (E_AUXILIAIRE.Text = '')) then exit;

{$IFDEF EAGLCLIENT}
    // TODO
{$ELSE}
    StConf := SQLConf('TIERS');
    if (StConf <> '') then StConf := ' AND ' + StConf;
{$ENDIF}
    if (Suiv) then Q := OpenSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE>"' + E_AUXILIAIRE.Text + '" ' + StConf + ' ORDER BY T_AUXILIAIRE',true)
    else Q := OpenSQL('Select T_AUXILIAIRE from TIERS where T_AUXILIAIRE<"' + E_AUXILIAIRE.Text + '" ' + StConf + ' order by T_AUXILIAIRE DESC', True);

    if (not Q.EOF) then
    begin
        E_AUXILIAIRE.Text := Q.Fields[0].AsString;
        AttribCollectif;
    end;

    Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 02/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.NextGene(Suiv:boolean);
var
    Q:TQuery;
    Existe:boolean;
    Nat:String3;
    StConf:string;
begin
    if (E_GENERAL.Text = '') then exit;

    Q := OpenSQL('Select G_NATUREGENE from GENERAUX where G_GENERAL="' + E_GENERAL.Text + '"', True);
    Existe := not Q.EOF;
    if (Existe) then
    begin
        Nat := Q.Fields[0].AsString;
        Ferme(Q);
    end
    else
    begin
        Ferme(Q);
        exit;
    end;

{$IFDEF EAGLCLIENT}
    // TODO
{$ELSE}
    StConf := SQLConf('GENERAUX');
    if (StConf <> '') then StConf := ' AND ' + StConf;
{$ENDIF}
    if (Suiv) then Q := OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL>"' + E_GENERAL.Text + '" AND G_NATUREGENE="' + Nat + '" ' + StConf + ' ORDER BY G_GENERAL',true)
    else Q := OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL<"' + E_GENERAL.Text + '" AND G_NATUREGENE="' + Nat + '" ' + StConf + ' ORDER BY G_GENERAL DESC',true) ;

    if (not Q.EOF) then E_GENERAL.Text := Q.Fields[0].AsString;
    Ferme(Q);

//    ControleFeuVert;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 06/02/2002
Modifié le ... :   /  /
Description .. : Nouvelle fonction de deletrage -> on envoie directement la
Suite ........ : requete sans ouvrir la fenetre
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.DeLettrage;
var
    i:integer;

    {***********A.G.L.***********************************************
    Auteur  ...... : Laurent Gendreau
    Créé le ...... : 14/10/2003
    Modifié le ... :   /  /
    Description .. : - LG 14/10/2003 - le champs e_paquetrevision n'etait pas
    Suite ........ : mis a jour
    Mots clefs ... :
    *****************************************************************}
    procedure Exec;
    var
        Query : string;
    begin
        Query := 'UPDATE ECRITURE SET E_DATEPAQUETMAX=E_DATECOMPTABLE,E_DATEPAQUETMIN=E_DATECOMPTABLE,E_LETTRAGE="", ' +
                 'E_ETATLETTRAGE="AL",E_LETTRAGEDEV="-",E_COUVERTURE=0,E_COUVERTUREDEV=0,E_REFLETTRAGE="", ' +
                 'E_ETAT="0000000000", E_TRESOSYNCHRO = "LET", '; {JP 26/04/04 : pour l'échéancier de la Tréso}
        Query := Query + 'E_PAQUETREVISION=1,';
        Query := Query + 'E_DATEMODIF="' + UsTime(NowH) + '" WHERE E_LETTRAGE="' + GL.Cells[9,GL.Row] + '" AND E_GENERAL="' + GL.Cells[1,GL.Row] + '" AND E_AUXILIAIRE="' + GL.Cells[2,GL.Row] + '"';
        ExecuteSQL(Query);
    end;

begin
 //LG* 05/01/2002
    try
        if ((GL.AllSelected) or (GL.nbSelected <> 0)) then
        begin
            if (GL.AllSelected) then
            begin
// BPY le 21/01/2004 => point 106 de la FFF : On as pas besoin et surtout on n'utilise pas QL dans Exec ... donc utilisation du GL ...
//                QL.First;
//                while (not QL.Eof) do
//                begin
//                    Exec;
//                    QL.Next;
//                end;
                for i := 1 to GL.RowCount-1 do
                begin
                    GL.Row := i;
                    Exec;
                end;
// fin BPY
            end
            else // Traitement <> si AllSelected car sinon ça merde, Si AllSelected alors NbSelected est faux, il vaut 0
            begin
                for i := 0 to GL.nbSelected - 1 do
                begin
                    GL.GotoLeBookMark(i);
                    // QL.Seek(GL.Row-1); => Seulement si on utilisait la Query a la place du grid ...
                    Exec;
                end;
            end;
        end; //if
    except
        on E:Exception do
        begin
            MessageAlerte('Erreur lors du delettrage. Toutes les opérations ont été annulées !' + #10#13#10#13 + E.message);
        end; // on
    end; // try
end;

{***********A.G.L.***********************************************
Auteur  ...... : VLA
Créé le ...... : ??/??/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.SetCritModified(Value: Boolean);
begin
    FCritModified := value;

    if (Value) then BChercher.Glyph := iCritGlyphModified.Picture.BitMap
    else BChercher.Glyph := iCritGlyph.Picture.BitMap;
end;

{***********A.G.L.***********************************************
Auteur  ...... : VLA
Créé le ...... : ??/??/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.CritereChange(Sender: TObject);
begin
  if (Loading) then exit;
  CritModified := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : VLA
Créé le ...... : ??/??/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulDLettre.Loaded;
var
    i : Integer;
begin
    inherited;

    if (V_PGI.AglDesigning) then exit;

    for i := 0 to ComponentCount - 1 do
    begin
        if (Components[i] is TControl) and (TControl(Components[i]).Parent is TTabSheet) then
        begin
            if (Components[i] is TControl) and TControl(Components[i]).Visible and (TControl(Components[i]).Enabled) then
            begin
                if ((Components[i] is THCritMaskEdit) and (not Assigned(THCritMaskEdit(Components[i]).OnEnter))) then THCritMaskEdit(Components[i]).OnEnter := V_PGI.EgaliseOnEnter;
                if ((Components[i] is TEdit) and (not assigned(TEdit(Components[i]).OnChange))) then TEdit(Components[i]).OnChange := CritereChange;
                if ((Components[i] is THValComboBox) and (not assigned(THValComboBox(Components[i]).OnClick))) then THValComboBox(Components[i]).OnClick := CritereChange;
                if ((Components[i] is TCheckBox) and (not assigned(TCheckBox(Components[i]).OnClick))) then TCheckBox(Components[i]).OnClick := CritereChange;
                if ((Components[i] is THCritMaskEdit) and (not assigned(THCritMaskEdit(Components[i]).OnChange))) then THCritMaskEdit(Components[i]).OnChange := CritereChange;
                if ((Components[i] is THRadioGroup) and (not assigned(THRadioGroup(Components[i]).OnClick))) then THRadioGroup(Components[i]).OnClick := CritereChange;
            end;
        end;
    end;
end;

end.
