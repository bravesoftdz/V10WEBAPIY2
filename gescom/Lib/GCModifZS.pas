unit GCModifZS;

interface

uses HSysMenu, HTB97,HEnt1, hmsgbox, Grids, ExtCtrls,Buttons,Hctrls,
     {$IFNDEF SANSCOMPTA}
     HCompte,
     {$ENDIF}
{$IFDEF EAGLCLIENT}
     UTob,
{$IFDEF GRC}
     MaineAGL,
{$ENDIF}
{$ELSE}
{$IFDEF GRC}
    Utob,FE_Main,
{$ENDIF}
    {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF GRC}
    UtilRT,EntRT,
{$ENDIF}
     StdCtrls, Forms,Ent1, Controls, Classes,Graphics,SysUtils, EntGC,
  TntButtons;

type
  TGCFModifZS = class(TForm)
    Panel1: TPanel;
    BAjouter: TToolbarButton97;
    BEnlever: TToolbarButton97;
    Panel2: TPanel;
    BOK: THBitBtn;
    BAnnuler: THBitBtn;
    BAide: THBitBtn;
    FZModifiable: TGroupBox;
    FListe: TStringGrid;
    FModif: TGroupBox;
    FBox: TScrollBox;
    Panels_On: TPanel;
    BValider: THBitBtn;
    MsgBox: THMsgBox;
    HMTrad: THSystemMenu;
    procedure FormShow(Sender: TObject);
    procedure BAjouterClick(Sender: TObject);
    procedure BEnleverClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure Panels_OnEnter(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BOKClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure Controls_OnExit(Sender: TObject);
  private    { Déclarations privées }
    FTable,FNatureTiers,SQLUpdate : string;
    CurLig            : LongInt;
    TitreFenetre : string;
    //CurTab : TTabOrder;
    OkRafale          : boolean;
    RepValide : integer;
    StockCtrl:string ;     //mcd 25/01/01 pour stocker le type d'enrgt à rechercher
    // Modif BTP
    Contexte : String;
{$IFDEF GRC}
    TobCl : Tob;
{$ENDIF}
    // -----
    procedure ChargeListe ;
    function  ConstruitUpdate : string;
    function  AffecteDataType(NomChamp : string) : String;
    procedure NewPosition;
    // Modif BTP
    Function ChampDansContexte(StTable: String; STChamp: String) : boolean;
    function champarticleCtxBtp(StChamp: String): boolean;
    function GetLibelleChampsCtx (Nomchamps,LibChamps : string) : string;
procedure TexteClick(Sender: TObject);
procedure TestExiste(Sender: TObject);
    // ------------
  public
    { Déclarations publiques }
  end;

function ModifZoneSerie(Table,Titre,NatureTiers : string; var Rafale:boolean ;var ReponseRafale:boolean;Contexte:String=''): string;

implementation
uses  CbpMCD
  		,CbpEnumerator
      ;


{$R *.DFM}

function ModifZoneSerie(Table,Titre,NatureTiers : string; var Rafale:boolean ;var ReponseRafale:boolean;Contexte:String=''): string;
var GCFModifZS : TGCFModifZS;
BEGIN
if Blocage(['nrCloture'],False,'nrAucun') then Exit ;
Result:='';
GCFModifZS:=TGCFModifZS.Create(Application);
try
  GCFModifZS.FTable:=Table;
  GCFModifZS.FNatureTiers:=NatureTiers;
  GCFModifZS.TitreFenetre:=Titre;
  //
  // Modif BTP
  //
  if Contexte <> '' then GCFModifZS.Contexte := Contexte;
  // ------

  //modif GA
  if Rafale then GCFModifZS.BOK.visible := false;
   
  if GCFModifZS.ShowModal=mrok then
     BEGIN
     ReponseRafale := GCFModifZS.RepValide=MrYes;
     Result:=GCFModifZS.SQLUpdate;
     Rafale:=GCFModifZS.OkRafale;
     END;
finally
  GCFModifZS.Free;
end;
END;

procedure TGCFModifZS.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
SQLUpDate:=''; Panels_On.Visible:=false;
Caption:=TitreFenetre;
ChargeListe;
UpdateCaption(Self) ;
end;

procedure TGCFModifZS.ChargeListe ;
Var St, stTable, stUnion, Ctrl : String ;
    R,iTable,iChamp  : integer;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
    NomChamps : string;
    LibChamps : string;
    TypeChamps : string;
BEGIN
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();

if FTable='' then Exit ;
stUnion:='';
//{$IFDEF EAGLCLIENT}
//{AFAIREEAGL}
//{$ELSE}
R:=0;
FListe.ColWidths[0]:=0; FListe.ColWidths[1]:=0;
If V_PGI.ExtendedFieldSelection<>'' then Ctrl:=V_PGI.ExtendedFieldSelection else Ctrl := 'Z';
StockCtrl:=V_PGI.ExtendedFieldSelection;  //mcd 25/01/01
// GRC V_PGI.ExtendedFieldSelection:='' ;
stTable:=ReadTokenSt(fTable);
while stTable<>'' do
   begin
   	table := Mcd.getTable(stTable);
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
       begin
       NomChamps := (FieldList.Current as IFieldCOM).name;
			 LibChamps := (FieldList.Current as IFieldCOM).libelle;
			 TypeChamps := (FieldList.Current as IFieldCOM).Tipe;

       //       if pos( ctrl,V_PGI.DEChamps[iTable,iChamp].Control)=0 then continue ;
       if LibChamps <> '' then // DBR Fiche 10093 au cas ou le libelle n'est pas chargé
                                                           // dans V_PGI.DECHamps, on ne peut pas tester la position 1
        if LibChamps[1]='.' then continue ; // pas les champs libres si le nom commence par un point
{$IFDEF GESCOM}
       if copy(NomChamps,1,12)='GPK_TOXAPPEL' then continue ; // pas les champs D'heure d'appel transmission caisse
{$ENDIF}
       // Modif BTP
       if (Contexte <> '') and (not ChampDansContexte (StTable,NomChamps)) then Continue;
       //
       FListe.RowCount:=R+1; FListe.Row:=R;
       St:='';
       FListe.Cells[0,R]:=NomChamps;
       FListe.Cells[1,R]:=TypeChamps;
       if ( copy(NomChamps,1,19) = 'YTC_TABLELIBRETIERS' ) or
          ( copy(NomChamps,1,17) = 'YTC_TABLELIBREFOU' ) then
          FListe.Cells[1,R]:='COMBO';
       if pos('LIBMUL',FListe.Cells[0,R])>0 then FListe.Cells[1,R]:='MULTI';
       FListe.Cells[2,R]:=GetLibelleChampsCtx(Nomchamps,LibChamps);
       if FListe.Cells[2,R]='' then FListe.Cells[2,R]:=NomChamps;
       inc(R);
       end;
   stTable:=ReadTokenSt(fTable);
   end;
FListe.Row:=0; FListe.SetFocus ;
//{$ENDIF}
END ;

{$IFDEF BTP}
Function TGCFModifZS.champarticleCtxBtp (StChamp: String): boolean;
var Autorise: string;
begin
Autorise:='';
result:=false;
// Liste des zones Interdites en fonction des contextes
if  Contexte = 'POU' then
    begin
    Autorise := 'GA_FAMILLENIV1;GA_FAMILLENIV2;GA_FAMILLENIV3;GA_COMPTAARTICLE;'+
                'GA_DPA;GA_DPR;GA_PVHT,GA_PVTTC;GA_PAHT;GA_PMAP;GA_PMRP;GA_FERME;'+
                'GA_QUALIFUNITEVTE;GA_QUALIFUNITESTO;GA_TARIFARTICLE;GA_TENUESTOCK;'+
                'GA_PCB;GA_FAMILLETAXE1;GA_VALLIBRE1;GA_VALLIBRE2;GA_VALLIBRE3;'+
                'GA_DATELIBRE1;GA_DATELIBRE2;GA_DATELIBRE3;GA_LIBREART1;GA_LIBREART2;'+
                'GA_LIBREART3;GA_LIBREART4;GA_LIBREART5;GA_LIBREART6;GA_LIBREART7;GA_LIBREART8;'+
                'GA_LIBREART9;GA_LIBREARTA;GA_CHARLIBRE1;GA_CHARLIBRE2;GA_CHARLIBRE3;GA_BOOLLIBRE1;'+
                'GA_BOOLLIBRE2;GA_BOOLLIBRE3;GA_REMISEPIED;GA_ESCOMPTABLE;GA_COMMISSIONNABLE;'+
                'GA_COEFFG;GA_CALCPRIXHT;GA_CALCPRIXTTC;GA_COEFCALCHT;GA_COEFCALCTTC;'+
                'GA_CALCAUTOHT;GA_CALCAUTOTTC;GA_FOURNPRINC;GA_DPRAUTO;GA_DOMAINE;GA_CALCPRIXPR;GA_NATUREPRES;'+
                'GA_PRIXPASMODIF;GA_QUALIFUNITEACH;';
    // Article de type pourcentage
    end else
if Contexte = 'OUV' then
    begin
    Autorise := 'GA_FAMILLENIV1;GA_FAMILLENIV2;GA_FAMILLENIV3;'+
                'GA_FERME;'+
                'GA_QUALIFUNITEVTE;'+
                'GA_VALLIBRE1;GA_VALLIBRE2;GA_VALLIBRE3;'+
                'GA_DATELIBRE1;GA_DATELIBRE2;GA_DATELIBRE3;GA_LIBREART1;GA_LIBREART2;'+
                'GA_LIBREART3;GA_LIBREART4;GA_LIBREART5;GA_LIBREART6;GA_LIBREART7;GA_LIBREART8;'+
                'GA_LIBREART9;GA_LIBREARTA;GA_CHARLIBRE1;GA_CHARLIBRE2;GA_CHARLIBRE3;GA_BOOLLIBRE1;'+
                'GA_BOOLLIBRE2;GA_BOOLLIBRE3;GA_REMISEPIED;GA_ESCOMPTABLE;GA_COMMISSIONNABLE;'+
                'GA_DOMAINE;GA_TYPENOMENC;';
    // Ouvrage ventilation au détail
    end else
if Contexte = 'OU1' then
    begin
    Autorise := 'GA_FAMILLENIV1;GA_FAMILLENIV2;GA_FAMILLENIV3;GA_COMPTAARTICLE;'+
                'GA_FERME;'+
                'GA_QUALIFUNITEVTE;'+
                'GA_FAMILLETAXE1;GA_VALLIBRE1;GA_VALLIBRE2;GA_VALLIBRE3;'+
                'GA_DATELIBRE1;GA_DATELIBRE2;GA_DATELIBRE3;GA_LIBREART1;GA_LIBREART2;'+
                'GA_LIBREART3;GA_LIBREART4;GA_LIBREART5;GA_LIBREART6;GA_LIBREART7;GA_LIBREART8;'+
                'GA_LIBREART9;GA_LIBREARTA;GA_CHARLIBRE1;GA_CHARLIBRE2;GA_CHARLIBRE3;GA_BOOLLIBRE1;'+
                'GA_BOOLLIBRE2;GA_BOOLLIBRE3;GA_REMISEPIED;GA_ESCOMPTABLE;GA_COMMISSIONNABLE;'+
                'GA_DOMAINE;GA_TYPENOMENC;';
    // Ouvrage ventilation en entete
    end else
if Contexte = 'MAC' then
    begin
    // Macro commande de saisie
    end else
if Contexte = 'MAR' then
    begin
    Autorise := 'GA_FAMILLENIV1;GA_FAMILLENIV2;GA_FAMILLENIV3;GA_COMPTAARTICLE;'+
                'GA_DPA;GA_DPR;GA_PVHT,GA_PVTTC;GA_PAHT;GA_PMAP;GA_PMRP;GA_FERME;'+
                'GA_QUALIFUNITEVTE;GA_QUALIFUNITESTO;GA_TARIFARTICLE;GA_TENUESTOCK;'+
                'GA_PCB;GA_FAMILLETAXE1;GA_VALLIBRE1;GA_VALLIBRE2;GA_VALLIBRE3;'+
                'GA_DATELIBRE1;GA_DATELIBRE2;GA_DATELIBRE3;GA_LIBREART1;GA_LIBREART2;'+
                'GA_LIBREART3;GA_LIBREART4;GA_LIBREART5;GA_LIBREART6;GA_LIBREART7;GA_LIBREART8;'+
                'GA_LIBREART9;GA_LIBREARTA;GA_CHARLIBRE1;GA_CHARLIBRE2;GA_CHARLIBRE3;GA_BOOLLIBRE1;'+
                'GA_BOOLLIBRE2;GA_BOOLLIBRE3;GA_REMISEPIED;GA_ESCOMPTABLE;GA_COMMISSIONNABLE;'+
                'GA_COEFFG;GA_CALCPRIXHT;GA_CALCPRIXTTC;GA_COEFCALCHT;GA_COEFCALCTTC;'+
                'GA_CALCAUTOHT;GA_CALCAUTOTTC;GA_FOURNPRINC;GA_DPRAUTO;GA_CALCPRIXPR;GA_NATUREPRES;'+
                'GA_PRIXPASMODIF;GA_DOMAINE;GA_GEREDEMPRIX;GA_FANTOME;GA_REMISELIGNE;GA_GERESAV;GA_QUALIFUNITEACH;';
    end else
if Contexte = 'ARP' then
    begin
    Autorise := 'GA_FAMILLENIV1;GA_FAMILLENIV2;GA_FAMILLENIV3;GA_COMPTAARTICLE;'+
                'GA_DPA;GA_DPR;GA_PVHT,GA_PVTTC;GA_PAHT;GA_PMAP;GA_PMRP;GA_FERME;'+
                'GA_QUALIFUNITEVTE;GA_QUALIFUNITESTO;GA_TARIFARTICLE;GA_TENUESTOCK;'+
                'GA_PCB;GA_FAMILLETAXE1;GA_VALLIBRE1;GA_VALLIBRE2;GA_VALLIBRE3;'+
                'GA_DATELIBRE1;GA_DATELIBRE2;GA_DATELIBRE3;GA_LIBREART1;GA_LIBREART2;'+
                'GA_LIBREART3;GA_LIBREART4;GA_LIBREART5;GA_LIBREART6;GA_LIBREART7;GA_LIBREART8;'+
                'GA_LIBREART9;GA_LIBREARTA;GA_CHARLIBRE1;GA_CHARLIBRE2;GA_CHARLIBRE3;GA_BOOLLIBRE1;'+
                'GA_BOOLLIBRE2;GA_BOOLLIBRE3;GA_REMISEPIED;GA_ESCOMPTABLE;GA_COMMISSIONNABLE;'+
                'GA_COEFFG;GA_CALCPRIXHT;GA_CALCPRIXTTC;GA_COEFCALCHT;GA_COEFCALCTTC;'+
                'GA_CALCAUTOHT;GA_CALCAUTOTTC;GA_FOURNPRINC;GA_DPRAUTO;GA_DOMAINE;GA_CALCPRIXPR;GA_NATUREPRES;'+
                'GA_PRIXPASMODIF;GA_GEREDEMPRIX;GA_FANTOME;GA_REMISELIGNE;GA_GERESAV;GA_QUALIFUNITEACH;';
    end else
if Contexte = 'PRE' then
    begin
    Autorise := 'GA_FAMILLENIV1;GA_FAMILLENIV2;GA_FAMILLENIV3;GA_COMPTAARTICLE;'+
                'GA_DPA;GA_DPR;GA_PVHT,GA_PVTTC;GA_PAHT;GA_PMAP;GA_PMRP;GA_FERME;'+
                'GA_QUALIFUNITEVTE;GA_TARIFARTICLE;'+
                'GA_FAMILLETAXE1;GA_VALLIBRE1;GA_VALLIBRE2;GA_VALLIBRE3;'+
                'GA_DATELIBRE1;GA_DATELIBRE2;GA_DATELIBRE3;GA_LIBREART1;GA_LIBREART2;'+
                'GA_LIBREART3;GA_LIBREART4;GA_LIBREART5;GA_LIBREART6;GA_LIBREART7;GA_LIBREART8;'+
                'GA_LIBREART9;GA_LIBREARTA;GA_CHARLIBRE1;GA_CHARLIBRE2;GA_CHARLIBRE3;GA_BOOLLIBRE1;'+
                'GA_BOOLLIBRE2;GA_BOOLLIBRE3;GA_REMISEPIED;GA_ESCOMPTABLE;GA_COMMISSIONNABLE;'+
                'GA_COEFFG;GA_CALCPRIXHT;GA_CALCPRIXTTC;GA_COEFCALCHT;GA_COEFCALCTTC;'+
                'GA_CALCAUTOHT;GA_CALCAUTOTTC;GA_DPRAUTO;GA_CALCPRIXPR;GA_NATUREPRES;'+
                'GA_PRIXPASMODIF;GA_DOMAINE;GA_GEREDEMPRIX;GA_FANTOME;GA_GEREANAL;GA_SECTION;GA_REMISELIGNE;GA_GERESAV;GA_FONCTION;';
    end else
if Contexte = 'FRA' then
    begin
    Autorise := 'GA_FAMILLENIV1;GA_FAMILLENIV2;GA_FAMILLENIV3;GA_COMPTAARTICLE;'+
                'GA_DPA;GA_DPR;GA_PVHT,GA_PVTTC;GA_PAHT;GA_PMAP;GA_PMRP;GA_FERME;'+
                'GA_QUALIFUNITEVTE;GA_TARIFARTICLE;'+
                'GA_FAMILLETAXE1;GA_VALLIBRE1;GA_VALLIBRE2;GA_VALLIBRE3;'+
                'GA_DATELIBRE1;GA_DATELIBRE2;GA_DATELIBRE3;GA_LIBREART1;GA_LIBREART2;'+
                'GA_LIBREART3;GA_LIBREART4;GA_LIBREART5;GA_LIBREART6;GA_LIBREART7;GA_LIBREART8;'+
                'GA_LIBREART9;GA_LIBREARTA;GA_CHARLIBRE1;GA_CHARLIBRE2;GA_CHARLIBRE3;GA_BOOLLIBRE1;'+
                'GA_BOOLLIBRE2;GA_BOOLLIBRE3;GA_REMISEPIED;GA_ESCOMPTABLE;GA_COMMISSIONNABLE;'+
                'GA_COEFFG;GA_CALCPRIXHT;GA_CALCPRIXTTC;GA_COEFCALCHT;GA_COEFCALCTTC;'+
                'GA_CALCAUTOHT;GA_CALCAUTOTTC;GA_DPRAUTO;GA_CALCPRIXPR;GA_NATUREPRES;'+
                'GA_PRIXPASMODIF;GA_DOMAINE;GA_GEREDEMPRIX;GA_FANTOME;GA_GEREANAL;GA_SECTION;GA_GERESAV;';
    end;
// Gestion BTP Obligatoire du fait que des zone ne sont pas visibles dans les fiches
if (Pos (stChamp,Autorise) > 0) then result := true ;
end;
{$ENDIF}

Function TGCFModifZS.GetLibelleChampsCtx (Nomchamps,LibChamps : string) : string;
begin
  if NomChamps  = 'GA_PRIXPASMODIF' then
  begin
    result := 'Application systématique des coefficients';
  end else
  begin
    result := LibChamps;
  end;
end;

Function TGCFModifZS.ChampDansContexte (StTable: String ;StChamp: String): boolean;
begin
result := true;
{$IFDEF BTP}
if StTable = 'ARTICLE' then result := champarticleCtxBtp (stchamp);
{$ENDIF}
end;

procedure TGCFModifZS.BAjouterClick(Sender: TObject);
var P  : TPanel;
    L  : TLabel;
    CC : TCheckBox;
    C  : THValComboBox;
    E  : THEdit;
    MC : THMultiValComboBox;
    TC : String;
    OkCase,OkCombo,OkEdit,OkMulti : boolean;
    i  : integer;
begin
if CurLig<0 then Exit ;
CurLig:=FListe.Row ;
TC:=FListe.Cells[1,CurLig] ;
if TC='' then Exit ;
OkCase:=(TC='BOOLEAN'); OkCombo:=(TC='COMBO'); OkMulti:=(TC='MULTI');
OkEdit:=((TC<>'BOOLEAN')And(TC<>'COMBO')And(TC<>'MULTI'));
if ((Not OkCase)And(Not OkCombo)And(Not OkEdit)And(Not OkMulti)) then Exit;  
P:=TPanel.Create(Self); P.Parent:=FBox; P.Name:='P'+IntToStr(CurLig);
P.ParentFont:=False; P.Font.Color:=clWindowText;
P.Height:=30; P.Align:=alTop; P.BevelInner:=bvLowered; P.BevelOuter:=bvRaised;
P.Caption:=''; P.OnEnter:=Panels_OnEnter; P.OnClick:=Panels_OnEnter;

if OkCase then
   begin
   CC:=TCheckBox.Create(Self); CC.Parent:=P; CC.Name:=FListe.Cells[0,CurLig];
   CC.Left:=8; CC.Width:=300; CC.Alignment:= taLeftJustify; CC.Top:=4;
   CC.Caption:=FListe.Cells[2,CurLig];
   //CC.OnExit:=Controls_OnExit; {JD}
   CC.setfocus;
   end
else
   begin
   L:=THLabel.Create(Self); L.Parent:=P; L.Name:='L'+IntToStr(CurLig);
   L.Left:=8; L.Width:=150; L.Top:=8; L.AutoSize:=false;
   L.Caption:=FListe.Cells[2,CurLig];
   if OkCombo then
      begin
      C:=THValComboBox.Create(Self); C.Parent:=P; C.Name:=FListe.Cells[0,CurLig];
      C.Left:=L.Left+L.Width+1; C.Width:=P.Width-C.Left-8; C.Top:=4; C.Text:='' ;
      C.Style:=csDropDownList ;
      C.ItemIndex:=0; L.FocusControl:=C;
      C.Vide:=True;
      C.VideString:='<<Aucun>>';
      C.DataType:=AffecteDataType(C.Name);
      //C.OnExit:=Controls_OnExit; {JD}
      C.SetFocus;
      end;
   if OkEdit then
      begin
      E:=THEdit.Create(Self); E.Parent:=P; E.Name:=FListe.Cells[0,CurLig];
      if TC='DATE' then begin E.EditMask:='##/##/####'; E.OpeType := otDate; end
      else if ((TC='INTEGER') or (TC='DOUBLE')) then E.OpeType := otReel
      else E.OpeType := otString;
      E.Left:=L.Left+L.Width+1; E.Width:=P.Width-E.Left-8; E.Top:=4;
      E.Text:=''; L.FocusControl:=E;
      E.DataType := AffecteDataType(E.Name);
      if E.DataType <> '' then E.ElipsisButton := True
      else E.ElipsisButton := False;
      //E.OnExit:=Controls_OnExit; {JD}
      E.SetFocus;
{$IFDEF GRC}
      if pos('RPR_RPRLIBTEXTE',E.Name) <> 0 then
         begin
         TobCl:=VH_RT.TobChampsPro.detail[0].FindFirst(['RDE_NOMCHAMP'],[E.Name],TRUE) ;
         if TobCl <> Nil then
            if TobCl.GetValue('RDE_TYPETEXTE') <> '' then
              begin
              E.ElipsisButton:=True;
              E.OnElipsisClick := TexteClick;
              E.OnExit := TestExiste;
              end;
         end;
{$ENDIF}
      end;
   if OkMulti then
      begin
      MC:=THMultiValComboBox.Create(Self); MC.Parent:=P; MC.Name:=FListe.Cells[0,CurLig];
      MC.Left:=L.Left+L.Width+1; MC.Width:=P.Width-MC.Left-8; MC.Top:=4; MC.Text:='' ;
      MC.Style:=csEditable ;
      MC.DataType:=AffecteDataType(MC.Name);
      MC.Complete:=True;
      L.FocusControl:=MC;
      MC.SetFocus;
      end;
   end;
FListe.RowHeights[CurLig]:=0;
NewPosition;
for i:=0 to FBox.ControlCount-1 do if FBox.Controls[i] is TPanel then
    BEGIN
    if TPanel(FBox.Controls[i])=P then TPanel(FBox.Controls[i]).Color:=clTeal
                                  else TPanel(FBox.Controls[i]).Color:=clBtnFace;
    END;
end;
{$IFDEF GRC}
procedure TGCFModifZS.TexteClick(Sender: TObject);
var Retour,StArg : String;
begin
TobCl:=VH_RT.TobChampsPro.detail[0].FindFirst(['RDE_NOMCHAMP'],[TEdit(Sender).Name],TRUE) ;
if TobCl = Nil then exit;
if TobCl.GetValue('RDE_TYPETEXTE') = '' then exit;

if  TobCl.GetValue('RDE_TYPETEXTE') = 'T' then
    begin
    ExecuteSQL('delete from filtres where FI_TABLE="RTMULTIERSCL" AND FI_LIBELLE="DEFAUT"');
    ExecuteSQL('INSERT INTO filtres (FI_TABLE,FI_LIBELLE,FI_CRITERES) SELECT FI_TABLE,"DEFAUT",FI_CRITERES FROM FILTRES WHERE FI_TABLE="RTMULTIERSCL" AND FI_LIBELLE="'+TobCl.GetValue('RDE_FILTRE')+'"');
    end;

if  TobCl.GetValue('RDE_TYPETEXTE') = 'R' then
    begin
    ExecuteSQL('delete from filtres where FI_TABLE="RTRESSOURCES" AND FI_LIBELLE="DEFAUT"');
    ExecuteSQL('INSERT INTO filtres (FI_TABLE,FI_LIBELLE,FI_CRITERES) SELECT FI_TABLE,"DEFAUT",FI_CRITERES FROM FILTRES WHERE FI_TABLE="RTRESSOURCES" AND FI_LIBELLE="'+TobCl.GetValue('RDE_FILTRE')+'"');
    end;

StArg:='';
if TobCl.GetValue('RDE_TYPETEXTE') = 'T' then
    begin
    if TEdit(Sender).Text <> '' then
       StArg:= 'T_TIERS='+ TEdit(Sender).Text;
    Retour:=AGLLanceFiche ('RT', 'RTTIERS_TL',StArg , '', 'PROSPECT');
    end
else
if TobCl.GetValue('RDE_TYPETEXTE') = 'R' then
    begin
    if TEdit(Sender).Text <> '' then
       StArg:= 'ARS_RESSOURCE='+ TEdit(Sender).Text;
    Retour:=AGLLanceFiche ('RT', 'RTRESSOURCE_TL',StArg , '', 'PROSPECT')
    end
;
if Retour <> '' then
   TEdit(Sender).Text:=ReadTokenSt(Retour);
end;

procedure TGCFModifZS.TestExiste(Sender: TObject);
var     Requete : String;
    Q : TQuery;
begin

if TobCl.GetValue('RDE_TYPETEXTE') = 'T' then
   Requete:='Select T_LIBELLE from TIERS where T_TIERS="'+THEdit(Sender).text+
   '" and (T_NATUREAUXI = "CLI" or T_NATUREAUXI="PRO")'
else
   if TobCl.GetValue('RDE_TYPETEXTE') = 'R' then
      Requete:='Select ARS_LIBELLE from RESSOURCE where ARS_RESSOURCE="'+THEdit(Sender).text+'"'
;
if Requete <> '' then
    begin
    Q:=OpenSQL(Requete,True,-1, '', True);
    if Q.Eof then
       begin
       PGIBox('Le code '+THEdit(Sender).Text+' n''existe pas','Code Inexistant');
       THEdit(Sender).SetFocus;
       end;
    Ferme (Q);
    end;
end;

{$ENDIF}
procedure TGCFModifZS.BEnleverClick(Sender: TObject);
var P : TPanel;
    T : TTabOrder;
    i : integer;
begin
P:= TPanel(FindComponent('P'+IntToStr(CurLig)));
if P=Nil then exit;
T:=P.TabOrder;
while P.ControlCount>0 do P.Controls[0].Free;
P.Free;
FListe.RowHeights[CurLig]:=FListe.DefaultRowHeight;
for i:=0 to FBox.ControlCount-1 do
    if FBox.Controls[i] is TPanel then
    begin
    if TPanel(FBox.Controls[i]).TabOrder=T-1 then TPanel(FBox.Controls[i]).Color:=clTeal
                                             else TPanel(FBox.Controls[i]).Color:=clBtnFace;
    end;
end;

procedure TGCFModifZS.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var P : TPanel;
begin
While FBox.ControlCount>0 do
   BEGIN
   P:= TPanel(FBox.Controls[0]);
   while P.ControlCount>0 do P.Controls[0].Free ;
   P.Free;
   END;
end;

procedure TGCFModifZS.Controls_OnExit(Sender: TObject);
var St     : string;
    Ctrl : TControl;
begin
St:=TComponent(Sender).Name;
Ctrl:=TControl(FindComponent(St));
if Ctrl=nil then exit;
if (Ctrl is THEdit) then
   begin
   //if THEdit(Ctrl).OpeType=OtDate then If Not IsValidDate(THEdit(Ctrl).Text) then THEdit(Ctrl).SetFocus;
   end;
end;

procedure TGCFModifZS.Panels_OnEnter(Sender: TObject);
var St     : string;
    i,IP,j : integer;
    P      : TPanel;
begin
St:=TComponent(Sender).Name; if Pos('P',St)<=0 then exit;
CurLig:=StrToInt(Copy(St,2,Length(St)));
for i:=0 to FBox.ControlCount-1 do
   if ((FBox.Controls[i] is TPanel) and (FBox.Controls[i].Name<>'Panels_On')) then
      BEGIN
      P:=TPanel(FBox.Controls[i]); St:=P.Name; if Pos('P',St)<=0 then break;
      IP:=StrToInt(Copy(St,2,Length(St)));
      if IP=CurLig then
         BEGIN
         P.Color:=clTeal;
         for j:=0 to P.ControlCount do
            BEGIN
            if P.Controls[j] is TCheckBox then BEGIN TCheckBox(P.Controls[j]).setfocus; break; END;
            if P.Controls[j] is THEdit then BEGIN THEdit(P.Controls[j]).setfocus; break; END;
            if P.Controls[j] is THValComboBox then BEGIN THValComboBox(P.Controls[j]).setfocus; break; END;
            if P.Controls[j] is THMultiValComboBox then BEGIN THMultiValComboBox(P.Controls[j]).setfocus; break; END;
            END;
         END else P.Color:=clBtnFace;
      END;
end;

procedure TGCFModifZS.FListeDblClick(Sender: TObject);
begin BAjouterClick(NIL); end;

function TGCFModifZS.ConstruitUpdate : string;
var i,j : integer;
    P   : TPanel;
    Champ,Valeur : String;
    C            : TControl;
    TSql : TQuery;
    stSql : string;
    BResult : boolean;
BEGIN
Result:='';
BResult := True;
For i:=0 to FBox.ControlCount-1 do
   BEGIN
   if (FBox.Controls[i] is TPanel) and (FBox.Controls[i].Name<>'Panels_On') then
      BEGIN
      P:=TPanel(FBox.Controls[i]);
      For j:=0 to P.ControlCount-1 do
          BEGIN
          C:=P.Controls[j];
          if C is TLabel then {RAF} else
             BEGIN
             Champ:=C.Name;
             if C is TCheckBox then
                BEGIN
                //if TCheckBox(C).state=cbGrayed then Valeur:='"0"'
                if TCheckBox(C).state=cbChecked then Valeur:='X'
                else if TCheckBox(C).state=cbUnChecked then Valeur:='-' else Valeur:=' ';
                END
             else if C is THValComboBox then
                BEGIN
                Valeur:=THValComboBox(C).Value;
                END
             else if C is THEdit then
                BEGIN
                Valeur:=THEdit(C).Text;
                if (Valeur <> '') and
                   ((Champ = 'T_PAYEUR') or (champ = 'T_FACTURE')) then
                    begin
                        stSql := 'SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="' + Valeur + '"';
                        if champ = 'T_PAYEUR' then
                            stSql := stSql + ' AND T_ISPAYEUR="X"';
                        TSql := OpenSql (stSql, True,-1, '', True);
                        if not TSql.Eof then
                            begin
                            Valeur := TSql.FindField ('T_AUXILIAIRE').asString;
                            end else BResult := False;
                        Ferme (TSql);
                    end;
                END
             else if C is THMultiValComboBox then
                BEGIN
                Valeur:=THMultiValComboBox(C).Text;
                END
     {$IFNDEF SANSCOMPTA}
             else if C is THCpteEdit then
                BEGIN
                Valeur:=THCpteEdit(C).Text;
                END ;
     {$ENDIF}
             END;
          END;
      if BResult then Result:=Result+Champ+'='+Valeur+'|' ;
      bResult := True;
      END;
   END;
END;


procedure TGCFModifZS.BValiderClick(Sender: TObject);
begin
if TBitBtn(Sender).Name='BValider' then
  begin
  {if (ctxGRC in V_PGI.PGIContexte) and (V_PGI.Superviseur = False) then
     begin
     PGIInfo('Vous n''avez pas les droits d''accès pour la modification en rafale','Modification en rafale');
     ModalResult:=mrNone;
     exit;
     end;}
  OkRafale:=True;
  RepValide:=PGIAskCancel('Voulez-vous effectuer la modification en série ?',Caption);
  if RepValide=mrCancel then begin ModalResult:=mrNone; exit; end;
  end
else OkRafale:=False ;
SQLUpdate:=ConstruitUpdate;
end;

procedure TGCFModifZS.BOKClick(Sender: TObject);
begin BValiderClick(Sender) ; end;

function TGCFModifZS.AffecteDataType(NomChamp : string) : String;
BEGIN
Result:='';
//{$IFDEF EAGLCLIENT}
//{AFAIREEAGL}
//{$ELSE}
  if NomChamp = 'GA_TYPENOMENC' Then
  begin
    Result := 'BTTYPENOMENC';
  end else
  if (NomChamp = 'GA_FAMILLENIV1') and (pos(contexte,'OUV;OU1')>0) Then
  begin
    Result := 'BTFAMILLEOUV1';
  end else
  if (NomChamp = 'GA_FAMILLENIV2') and (pos(contexte,'OUV;OU1')>0) Then
  begin
    Result := 'BTFAMILLEOUV2';
  end else
  if (NomChamp = 'GA_FAMILLENIV3') and (pos(contexte,'OUV;OU1')>0) Then
  begin
    Result := 'BTFAMILLEOUV3';
  end else
  if (NomChamp = 'GA_DOMAINE') Then
  begin
    Result := 'BTDOMAINEACT';
  end else
  begin
    Result:=Get_Join(NomChamp);
  end;
//{$ENDIF}
//if Result='' then Result:=AffectettdeChoixCod(NomChamp) ;
END;


procedure TGCFModifZS.NewPosition;
var i, New : integer;
BEGIN
New:=CurLig;
for i:=CurLig to FListe.RowCount-2 do
   if FListe.RowHeights[i]=0 then New:=New+1 else break;
if FListe.RowHeights[New]=0 then
   BEGIN
   New:=CurLig;
   for i:=CurLig downto 1 do
      if FListe.RowHeights[i]=0 then New:=New-1 else break;
   END;
if FListe.RowHeights[New]=0 then CurLig:=-1 else BEGIN CurLig:=New; FListe.Row:=New; END;
END;



procedure TGCFModifZS.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.

