unit AssistRecopiTarf;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
  HPanel,HEnt1,Ent1,
{$IFDEF EAGLCLIENT}
{$ELSE}
   dbTables,DBGrids, db,MenuOLG,
{$ENDIF}
     UTob,Grids,UDimArticle,Math,TarifUtil ;

  Procedure RecopieTarifMode ;

type
  TFRecopiTarf = class(TFAssist)
    TabSheet1: TTabSheet;
    PTITRE: THPanel;
    TINTRO: THLabel;
    GBTARIF: TGroupBox;
    TYPETARIF: THValComboBox;
    TTYPETARIF: THLabel;
    TPERIODE: THLabel;
    PERIODE: THValComboBox;
    TabSheet2: TTabSheet;
    GBTYPE: TGroupBox;
    CB_TARFCATART: TCheckBox;
    CB_TARFART: TCheckBox;
    CB_TARFCLI: TCheckBox;
    CB_TARFCATCLI: TCheckBox;
    GLISTE: THGrid;
    GCOPIE: THGrid;
    TLISTEET: THLabel;
    MAZE: TCheckBox;
    BFLECHEDROITE2: TToolbarButton97;
    BFLECHEGAUCHE2: TToolbarButton97;
    TabSheet3: TTabSheet;
    GBPRIX: TGroupBox;
    GBREMISE: TGroupBox;
    COEF: THNumEdit;
    ARRONDIP: THValComboBox;
    TRECOPIER: THLabel;
    TDANS: THLabel;
    TCOEF: THLabel;
    TARRONDIP: THLabel;
    TMONTANT: THLabel;
    MONTANT: THNumEdit;
    HMsgErr: THMsgBox;
    TabSheet4: TTabSheet;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    TRecap: THLabel;
    ListRecap: TListBox;
    HRecap: THMsgBox;
    BFinRcap: TButton;
    Bevel1: TBevel;
    Bevel2: TBevel;
    Bevel3: TBevel;
    Bevel4: TBevel;
    BFLECHETOUS: TToolbarButton97;
    BFLECHEAUCUN: TToolbarButton97;
    ARECOPIER: THValComboBox;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure bFinClick(Sender: TObject) ;

    // Onglet 1
    Procedure TYPETARIFOnChange (Sender: TObject) ;
    Procedure PERIODEOnChange (Sender: TObject) ;
    Procedure CheckBoxOnClick (Sender: TObject) ;

    // Onglet 2
    procedure ClickFlecheDroite2(Sender: TObject);
    procedure ClickFlecheTous(Sender: TObject);
    procedure ClickFlecheGauche2(Sender: TObject);
    procedure ClickFlecheAucun(Sender: TObject) ;

    // Onglet 4
    procedure ClickBFinRecap(Sender: TObject) ;

  private
     TobListe,TobCopie,TobTarifSource,TobTarifDest : TOB ;
     NbTarifCreer: Integer ;
    procedure RemplirGridEtablissement ;
    procedure RefreshGrid(posListe,posCopie : integer) ;
    procedure RefreshBouton ;
    function  SpaceStr ( nb : integer) : string;
    function  ExtractLibelle ( St : string) : string;
    procedure ListeRecap;
    procedure ListeRecapTermine;

    // Traitement
    procedure RecopieTarif ;
    Procedure CreerTobTarifSource ;
    procedure CopieTarif ;
    procedure TraiterPrix(TOBD: Tob) ;
    function  PrixContre(Prix:Double; Devise:String):Double ;
    function  CreerTobDepot(Depot:String):TOB ;
    Function  WhereSQL (Art,Cli,CatArt,CatCli: Boolean) : String ;
    procedure SupprimerTarifDepot ;

    procedure LibereTout ;

  public
    { Déclarations publiques }

  end;

var
  FRecopiTarf: TFRecopiTarf;
  i_NumEcran : integer;

Function VerifTarifExistant(TobDepot,TobTarifFille:TOB): TOB ;

implementation

{$R *.DFM}
Procedure RecopieTarifMode ;
var
  FRecopiTarf: TFRecopiTarf;
BEGIN
	FRecopiTarf:=TFRecopiTarf.Create(Application) ;
   Try
       FRecopiTarf.ShowModal;
   Finally
       FRecopiTarf.free;
   End;
END ;

{=========================================================================================}
{============================= Evenements de la forme ====================================}
{=========================================================================================}
procedure TFRecopiTarf.FormShow(Sender: TObject);
var Onglet : TTabSheet;
    St_NomOnglet : String;
begin
  inherited;
bAnnuler.Visible := True;
bSuivant.Enabled := False;
bFin.Visible := True;
bFin.Enabled := False;
BFinRcap.Visible:=False ;
TYPETARIF.Plus:='(GFT_CODETYPE<>"...")' ;
i_NumEcran := 0;
Onglet := P.ActivePage;
st_NomOnglet := Onglet.Name;
i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
//RemplirGridEtablissement ;
//RefreshBouton ;
end;

procedure TFRecopiTarf.FormClose(Sender: TObject; var Action: TCloseAction);
begin
LibereTout ;
{$IFDEF EAGLCLIENT}
{$ELSE}
   FMenuG.VireInside(nil) ;
{$ENDIF}

end;


procedure TFRecopiTarf.bSuivantClick(Sender: TObject);
Var Onglet : TTabSheet;
    St_NomOnglet : String;
begin
  inherited;
Onglet := P.ActivePage;
st_NomOnglet := Onglet.Name;
i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
if (i_NumEcran = 2) then
   begin
   if ((ARECOPIER.Value='') or (TobCopie.Detail.Count=0)) then
     begin
     RestorePage ;
     if (ARECOPIER.Text='') then PGIBox (HMsgErr.Mess[0],Caption)
     else if (not TobCopie.Detail.Count>0) then PGIBox (HMsgErr.Mess[1],Caption) ;
     Onglet := PreviousPage ;
     if Onglet = nil then P.SelectNextPage(True) else BEGIN P.ActivePage := Onglet ; PChange(nil) ; END ;
     end ;
    end ;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
if bFin.Enabled then ListeRecap;
end ;

procedure TFRecopiTarf.bPrecedentClick(Sender: TObject);
Var Onglet : TTabSheet;
    St_NomOnglet : String;
begin
  inherited;
Onglet := P.ActivePage;
st_NomOnglet := Onglet.Name;
i_NumEcran := strtoint (Copy (st_NomOnglet, length (st_NomOnglet), 1)) - 1;
if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
end ;

procedure TFRecopiTarf.bFinClick(Sender: TObject);
begin
  inherited;
RecopieTarif ;
end;

{=========================================================================================}
{============================= Evenements Onglet 1 =======================================}
{=========================================================================================}
Procedure TFRecopiTarf.TYPETARIFOnChange (Sender: TObject) ;
 begin
   inherited;
   bSuivant.Enabled:=(TYPETARIF.Value<>'') and (PERIODE.Value<>'');
   ARECOPIER.Value:='' ;
   ARECOPIER.Plus:='(ET_TYPETARIF = "'+TYPETARIF.Value+'")';
   RemplirGridEtablissement ;
   RefreshBouton ;
 end ;

Procedure TFRecopiTarf.PERIODEOnChange (Sender: TObject) ;
 begin
   inherited;
   bSuivant.Enabled:=(TYPETARIF.Value<>'') and (PERIODE.Value<>'');
 end ;

Procedure TFRecopiTarf.CheckBoxOnClick (Sender: TObject) ;
begin
   inherited;
   bSuivant.Enabled:=not ((not CB_TARFART.Checked) And (not CB_TARFCATART.Checked) And (not CB_TARFCLI.Checked) And (not CB_TARFCATCLI.Checked)) ;
end ;

{=========================================================================================}
{============= Evenements de l'onglet 2 -consernant les établissement ====================}
{=========================================================================================}
procedure TFRecopiTarf.RemplirGridEtablissement ;
Var QQ:TQuery ;
j: Integer ;
begin
if TobListe=nil then TobListe:=TOB.CREATE('Les etablissements liste',NIL,-1) ;
if TobCopie=nil then TobCopie:=TOB.CREATE('Les etablissement copie',NIL,-1) ;
QQ:=OpenSQL('select ET_ETABLISSEMENT,ET_ABREGE from ETABLISS where ET_TYPETARIF="'+TYPETARIF.VALUE+'" order by ET_ETABLISSEMENT',True) ;
if not QQ.EOF then TobListe.LoadDetailDB('ETABLISS','','',QQ,False) ;
j:=TobListe.Detail.count ;
if TobListe.FindFirst(['ET_ETABLISSEMENT'],[''],True)=nil then
  begin
  TOB.Create ('ETABLISS', TobListe,j) ;
  TobListe.Detail[j].PutValue('ET_ETABLISSEMENT','') ;
  TobListe.Detail[j].PutValue('ET_ABREGE','Toutes boutiques') ;
  end;
TobListe.PutGridDetail(GLISTE,False,False,'ET_ETABLISSEMENT;ET_ABREGE',True) ;
GLISTE.ColWidths[0]:=30 ; GLISTE.ColAligns[0]:=taCenter ;
GLISTE.ColWidths[1]:=208 ; GLISTE.ColAligns[1]:=taLeftJustify ;
GCOPIE.ColWidths[0]:=30 ; GCOPIE.ColAligns[0]:=taCenter ;
GCOPIE.ColWidths[1]:=208 ; GCOPIE.ColAligns[1]:=taLeftJustify ;
GCOPIE.RowCount:=0 ;
Ferme(QQ) ;
end ;

procedure TFRecopiTarf.ClickFlecheDroite2 ;
var indiceFille : integer ;
begin
// Y a t il quelque chose de sélectionné ?
if GLISTE.Row<0 then exit ;
// Changement du parent de l'élément de la liste des établissements
if TobCopie.Detail.Count>0 then indiceFille:=GCOPIE.Row+1 else indiceFille:=0 ;
TobListe.detail[GLISTE.Row].ChangeParent(TobCopie,indiceFille) ;
RefreshGrid(GLISTE.Row,GCOPIE.Row+1) ;
end ;

procedure TFRecopiTarf.ClickFlecheTous ;
var indiceFille,iGrd,posListe : integer ;
begin
if GLISTE.Row<0 then exit ;
// Changement du parent de l'élément de la liste des établissements
if TobCopie.Detail.Count>0 then indiceFille:=GCOPIE.Row+1 else indiceFille:=0 ;
posListe:=TobListe.detail.count-1 ;
for iGrd:=0 to posListe do TobListe.detail[0].ChangeParent(TobCopie,indiceFille+iGrd) ;
RefreshGrid(0,indiceFille+posListe) ;
end ;

procedure TFRecopiTarf.ClickFlecheGauche2 ;
var indiceFille : integer ;
begin
// Y a t il quelque chose de sélectionné ?
if GCOPIE.Row<0 then exit ;
// Changement du parent de l'élément des établissements affichés
if TobListe.Detail.Count>0 then indiceFille:=GLISTE.Row+1 else indiceFille:=0 ;
TobCopie.detail[GCOPIE.Row].ChangeParent(TobListe,indiceFille) ;
RefreshGrid(GLISTE.Row+1,GCOPIE.Row) ;
end ;

procedure TFRecopiTarf.ClickFlecheAucun ;
var indiceFille,iGrd,posMAJ : integer ;
begin
if GCOPIE.Row<0 then exit ;
// Changement du parent de l'élément de la liste des établissements
if TobListe.Detail.Count>0 then indiceFille:=GLISTE.Row+1 else indiceFille:=0 ;
posMAJ:=TobCopie.detail.count-1 ;
for iGrd:=0 to posMAJ do TobCopie.detail[0].ChangeParent(TobListe,indiceFille+iGrd) ;
RefreshGrid(indiceFille+posMAJ,0) ;
end ;

procedure TFRecopiTarf.RefreshGrid(posListe,posCopie : integer) ;
begin
TobCopie.PutGridDetail(GCOPIE,False,False,'ET_ETABLISSEMENT;ET_ABREGE',True) ;
TobListe.PutGridDetail(GLISTE,False,False,'ET_ETABLISSEMENT;ET_ABREGE',True) ;
GCOPIE.Row:=Min(posCopie,GCOPIE.RowCount-1) ;
GLISTE.Row:=Min(posListe,GLISTE.RowCount-1) ;
RefreshBouton ;
end ;

procedure TFRecopiTarf.RefreshBouton ;
begin
BFLECHEDROITE2.Enabled := TobListe.Detail.Count>0 ;
BFLECHEGAUCHE2.Enabled := TobCopie.Detail.Count>0 ;
end ;

{=========================================================================================}
{================================= Récapitulatif =========================================}
{=========================================================================================}
function TFRecopiTarf.SpaceStr ( nb : integer) : string;
Var St_Chaine : string ;
    i_ind : integer ;
BEGIN
St_Chaine := '' ;
for i_ind := 1 to nb do St_Chaine:=St_chaine+' ';
Result:=St_Chaine;
END;

function TFRecopiTarf.ExtractLibelle ( St : string) : string;
Var St_Chaine : string ;
    i_pos : integer ;
BEGIN
Result := '';
i_pos := Pos ('&', St);
if i_pos > 0 then
    BEGIN
    St_Chaine := Copy (St, 1, i_pos - 1) + Copy (St, i_pos + 1, Length(St));
    END else St_Chaine := St;
Result := St_Chaine + ' : ';
END;

procedure TFRecopiTarf.ListeRecap;
Var st_chaine : string;
i: Integer ;
BEGIN
ListRecap.Items.Clear;
ListRecap.Items.Add (PTITRE.Caption);
ListRecap.Items.Add ('');
//Info concerant le premier onglet
ListRecap.Items.Add ('Tarifs :');
if TYPETARIF.Value<>'' then
    ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (TTYPETARIF.Caption) + TYPETARIF.Text);
if PERIODE.Value<>'' then
    ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (TPERIODE.Caption) + PERIODE.Text);
ListRecap.Items.Add (ExtractLibelle (GBTYPE.Caption));
if CB_TARFART.checked then
        BEGIN
        st_chaine := HRecap.Mess[0];
        ListRecap.Items.Add (SpaceStr(4) +ExtractLibelle (CB_TARFART.Caption) + st_chaine);
        END else
        begin
        st_chaine := HRecap.Mess[1];
        ListRecap.Items.Add (SpaceStr(4) +ExtractLibelle (CB_TARFART.Caption) + st_chaine);
        end ;
if CB_TARFCATART.checked then
        BEGIN
        st_chaine := HRecap.Mess[0];
        ListRecap.Items.Add (SpaceStr(4) +ExtractLibelle (CB_TARFCATART.Caption) + st_chaine);
        END else
        begin
        st_chaine := HRecap.Mess[1];
        ListRecap.Items.Add (SpaceStr(4) +ExtractLibelle (CB_TARFCATART.Caption) + st_chaine);
        end ;
if CB_TARFCLI.checked then
        BEGIN
        st_chaine := HRecap.Mess[0];
        ListRecap.Items.Add (SpaceStr(4) +ExtractLibelle (CB_TARFCLI.Caption) + st_chaine);
        END else
        begin
        st_chaine := HRecap.Mess[1];
        ListRecap.Items.Add (SpaceStr(4) +ExtractLibelle (CB_TARFCLI.Caption) + st_chaine);
        end ;
if CB_TARFCATCLI.checked then
        BEGIN
        st_chaine := HRecap.Mess[0];
        ListRecap.Items.Add (SpaceStr(4) +ExtractLibelle (CB_TARFCATCLI.Caption) + st_chaine);
        END else
        begin
        st_chaine := HRecap.Mess[1];
        ListRecap.Items.Add (SpaceStr(4) +ExtractLibelle (CB_TARFCATCLI.Caption) + st_chaine);
        end ;
ListRecap.Items.Add ('');
//Info concerant le deuxième onglet
if ARECOPIER.Value<>'' then
    ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (TRECOPIER.Caption) + ARECOPIER.Text);
if TobCopie.Detail.Count>0 then
   begin
    ListRecap.Items.Add ('Dans les établissement(s):' );
    For i:=0 to TobCopie.Detail.count-1 do
      begin
      ListRecap.Items.Add (SpaceStr(4) + TobCopie.Detail[i].GetValue('ET_ABREGE'));
      end ;
   end ;
if MAZE.checked then
        BEGIN
        st_chaine := HRecap.Mess[0];
        ListRecap.Items.Add (ExtractLibelle (MAZE.Caption) + st_chaine);
        END else
        begin
        st_chaine := HRecap.Mess[1];
        ListRecap.Items.Add (ExtractLibelle (MAZE.Caption) + st_chaine);
        end ;
ListRecap.Items.Add ('');
if MAZE.Checked=False then
  begin
  //Info concerant le troisième onglet
  ListRecap.Items.Add (ExtractLibelle (GBPRIX.Caption)) ;
  ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (TCOEF.Caption) + COEF.Text);
  ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (TARRONDIP.Caption) + ARRONDIP.Text);
  ListRecap.Items.Add (ExtractLibelle (GBREMISE.Caption)) ;
  ListRecap.Items.Add (SpaceStr(4) + ExtractLibelle (TMONTANT.Caption) + MONTANT.Text);
  ListRecap.Items.Add ('');
  end ;
END;

procedure TFRecopiTarf.ListeRecapTermine;
Var i: Integer ;
BEGIN
ListRecap.Items.Clear;
ListRecap.Items.Add ('La recopie des tarifs est terminée.');
// Etablissement à recopier
ListRecap.Items.Add(HRecap.Mess[5]) ;
ListRecap.Items.Add(SpaceStr(4)+ARECOPIER.Text) ;
ListRecap.Items.Add ('');
// Dans les établissements
ListRecap.Items.Add(HRecap.Mess[6]) ;
for i:=0 to TobCopie.Detail.count-1 do
  begin
  ListRecap.Items.Add (SpaceStr(4)+TobCopie.Detail[i].Getvalue('ET_ABREGE')) ;
  end ;
ListRecap.Items.Add ('') ;
// Nb enreg
ListRecap.Items.Add(IntToStr(TobTarifDest.Detail.count) +' '+ HRecap.Mess[4]) ;
BFinRcap.Visible:=True ;
END;

procedure TFRecopiTarf.ClickBFinRecap(Sender: TObject) ;
begin
close ;
RecopieTarifMode ;
end ;

procedure TFRecopiTarf.LibereTout ;
begin
TobListe.Free ; TobListe:=Nil ;
TobCopie.Free ; TobCopie:=Nil ;
TobTarifSource.Free ; TobTarifSource:=Nil ;
TobTarifDest.Free ; TobTarifDest:=Nil ;
end ;

{=========================================================================================}
{================================== Traitement ===========================================}
{=========================================================================================}
procedure TFRecopiTarf.RecopieTarif ;
Var ioerr : TIOErr ;
    TOBJnal : TOB ;
    QQ : TQuery ;
    NumEvt,NbEnreg : integer ;
BEGIN
CreerTobTarifSource ;
NbEnreg:=TOBCopie.detail.count ;
if NbEnreg <= 0 then exit;
bfin.Enabled:=False ;
NumEvt:=0 ;
TOBJnal:=TOB.Create('JNALEVENT', Nil, -1) ;
TOBJnal.PutValue('GEV_TYPEEVENT', 'TAR');
TOBJnal.PutValue('GEV_LIBELLE', PTITRE.Caption);
TOBJnal.PutValue('GEV_DATEEVENT', Date);
TOBJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
ioerr := Transactions (CopieTarif , 2);
QQ:=OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',True) ;
if Not QQ.EOF then NumEvt:=QQ.Fields[0].AsInteger ;
Ferme(QQ) ;
Inc(NumEvt) ;
TOBJnal.PutValue('GEV_NUMEVENT', NumEvt);
Case ioerr of
        oeOk  : BEGIN
                TOBJnal.PutValue('GEV_ETATEVENT', 'OK');
                TOBJnal.PutValue('GEV_BLOCNOTE', ListRecap.Items.Text);
                END;
    oeUnknown : BEGIN
                MessageAlerte(HRecap.Mess[2]) ;
                TOBJnal.PutValue('GEV_ETATEVENT', 'ERR');
                END ;
    oeSaisie  : BEGIN
                MessageAlerte(HRecap.Mess[3]) ;
                TOBJnal.PutValue('GEV_ETATEVENT', 'ERR');
                END ;
   END ;
TOBJnal.InsertDB(Nil) ;
TOBJnal.Free ;
ListeRecapTermine ;
END;

Procedure TFRecopiTarf.CreerTobTarifSource ;
var CodeDepot,SQL: String ;
QTarif: TQuery ;
TarifMode: Integer ;
TarfArt,TarfCatArt,TarfCli,TarfCatCli :Boolean ;
begin
TOBTarifSource:=TOB.Create('LesTarifs',nil,-1) ;
CodeDepot:=ARECOPIER.Value ;
TarifMode:=RecupCodeTarifMode(TYPETARIF.Value,PERIODE.Value,'VTE') ;
If CB_TARFART.Checked then TarfArt:=True else TarfArt:=False ;
If CB_TARFCATART.Checked then TarfCatArt:=True else TarfCatArt:=False ;
If CB_TARFCLI.Checked then TarfCli:=True else TarfCli:=False ;
If CB_TARFCATCLI.Checked then TarfCatCli:=True else TarfCatCli:=False ;
SQL:='Select * from tarif where ' ;
SQL:=SQL + WhereSQL(TarfArt,TarfCli,TarfCatArt,TarfCatCli) ;
SQL:=SQL + ') AND GF_DEPOT="'+CodeDepot+'"'+
         ' AND GF_TARFMODE="'+IntToStr(TarifMode)+'" AND GF_FERME="-"' ;
SQL:=SQL+' ORDER BY GF_PRIORITE DESC' ;
QTarif:=OpenSQL(SQL,True) ;
if Not QTarif.EOF then TOBTarifSource.LoadDetailDB('TARIF','','',QTarif,False)
  else begin
  PGIBox(HMsgErr.Mess[2],PTitre.Caption) ;
  end ;
Ferme(QTarif) ;
end ;

procedure TFRecopiTarf.CopieTarif ;
var i,j,k,MaxTarif: Integer ;
Q:TQuery ;
TobDepot,TobExist: Tob;
Info: String ;
begin
NbTarifCreer:=0 ;
if MAZE.Checked then SupprimerTarifDepot ;
k:=0 ;
Q := OpenSQL ('SELECT MAX(GF_TARIF) FROM TARIF', TRUE) ;
if Q.EOF then MaxTarif := 1 else MaxTarif := Q.Fields[0].AsInteger + 1 ;
Ferme(Q) ;
TobTarifDest:=TOB.Create('',nil,-1) ;
For i:=0 to TobCopie.Detail.count-1 do
  begin
  TobDepot:=CreerTobDepot(TobCopie.Detail[i].GetValue('ET_ETABLISSEMENT')) ;
  For j:=0 to TOBTarifSource.Detail.count-1 do
    begin
     if TobCopie.Detail[i].GetValue('ET_ETABLISSEMENT')<>TobTarifSource.Detail[j].GetValue('GF_DEPOT') then
       begin
       NbTarifCreer:=NbTarifCreer+1 ;
       TobExist:=VerifTarifExistant(TobDepot,TobTarifSource.Detail[j]) ;
       if TobExist <> nil then
         begin
         TOB.Create ('TARIF', TobTarifDest, k) ;
         TobTarifDest.Detail[k].Dupliquer(TobExist, False, True) ;
         TobTarifDest.Detail[k].PutValue('GF_PRIXUNITAIRE',TOBTarifSource.Detail[j].GetValue('GF_PRIXUNITAIRE')) ;
         TobTarifDest.Detail[k].PutValue('GF_CALCULREMISE',TOBTarifSource.Detail[j].GetValue('GF_CALCULREMISE')) ;
         TraiterPrix(TobTarifDest.Detail[k]) ;
         k:=k+1 ;
         end else
         begin
         TOB.Create ('TARIF', TobTarifDest, k) ;
         TobTarifDest.Detail[k].Dupliquer(TOBTarifSource.Detail[j], False, True);
         TobTarifDest.Detail[k].PutValue('GF_DEPOT',TobCopie.Detail[i].GetValue('ET_ETABLISSEMENT')) ;
         TobTarifDest.Detail[k].PutValue('GF_SOCIETE',TobCopie.Detail[i].GetValue('ET_ETABLISSEMENT')) ;
         TobTarifDest.Detail[k].PutValue('GF_TARIF',MaxTarif) ;
         Info:=RecupInfoPeriode (PERIODE.Value,TobCopie.Detail[i].GetValue('ET_ETABLISSEMENT')) ;
         TobTarifDest.Detail[k].PutValue('GF_DATEDEBUT',StrToDate(ReadTokenSt(Info))) ;
         TobTarifDest.Detail[k].PutValue('GF_DATEFIN',StrToDate(ReadTokenSt(Info))) ;
         TobTarifDest.Detail[k].PutValue('GF_ARRONDI',ReadTokenSt(Info)) ;
         TobTarifDest.Detail[k].PutValue('GF_CASCADEREMISE',ReadTokenSt(Info)) ;
         TobTarifDest.Detail[k].PutValue('GF_NATUREAUXI','CLI') ;
         TobTarifDest.Detail[k].PutValue('GF_DEMARQUE',ReadTokenSt(Info)) ;
         TraiterPrix(TobTarifDest.Detail[k]) ;
         MaxTarif:=MaxTarif+1 ;
         K:=k+1 ;
         end ;
       end ;
    end ;
    TOBDepot.Free; //TOBDepot:=nil ;
  end ;
TobTarifDest.InsertOrUpdateDB (True) ;
end ;

procedure TFRecopiTarf.SupprimerTarifDepot ;
var i,TarifMode:Integer ;
SQL: String ;
begin
TarifMode:=RecupCodeTarifMode(TYPETARIF.Value,PERIODE.Value,'VTE') ;
SQL:='DELETE FROM TARIF WHERE ' ;
SQL:=SQL + WhereSQL(CB_TARFART.Checked,CB_TARFCLI.Checked,CB_TARFCATART.Checked,CB_TARFCATCLI.Checked) ;
SQL:=SQL + ') AND GF_TARFMODE="'+IntToStr(TarifMode)+'" AND GF_FERME="-"' ;
SQL:=SQL + ' AND (GF_DEPOT="'+TobCopie.Detail[0].GetValue('ET_ETABLISSEMENT')+'"' ;
for i:=1 to TobCopie.Detail.count-1 do
  begin
  SQL:=SQL + ' OR GF_DEPOT="'+TobCopie.Detail[i].GetValue('ET_ETABLISSEMENT')+'"' ;
  //ExecuteSQL(SQL) ;
  end ;
SQl:=SQL + ')' ;
ExecuteSQL(SQL) ;
end ;

procedure TFRecopiTarf.TraiterPrix(TOBD: Tob) ;
var Prix,Remise,PrixCon: double ;
begin
Prix:=TOBD.GetValue('GF_PRIXUNITAIRE') ;
Remise:=Valeur(TOBD.GetValue('GF_CALCULREMISE')) ;
// Prix
if Prix <> 0 then
  begin
  if COEF.Value <> 0 then
    begin
    Prix:=Prix*COEF.Value ;
    Prix:=ArrondirPrix(ARRONDIP.Value,Prix) ;
    end else Prix:=0 ;
  TOBD.PutValue ('GF_PRIXUNITAIRE',Prix) ;
  TOBD.PutValue ('GF_ARRONDI',ARRONDIP.Value) ;
  end ;
// Remise
if Remise <> 0 then
  begin
  if MONTANT.Value <> 0 then
    begin
    Remise:=Remise+MONTANT.Value ;
    TOBD.PutValue ('GF_CALCULREMISE',FloatToStr(Remise)) ;
    TOBD.PutValue ('GF_REMISE', FloatToStr(Remise));
    //TOBD.PutValue ('GF_ARRONDI',ARRONDIR.Value) ;
    end ;
  end ;
PrixCon:=PrixContre(Prix, TOBD.GetValue('GF_DEVISE')) ;
end ;

function TFRecopiTarf.PrixContre(Prix:Double; Devise:String):Double ;
var PrixCon,PrixConAr: Double;
begin
if VH^.TenueEuro then ToxConvToDev2 (Prix,PrixCon,PrixConAr,Devise, V_PGI.DeviseFongible, Date, nil)
else ToxConvToDev2 (Prix,PrixCon,PrixConAr,Devise, 'EUR', Date, nil) ;
result:=PrixConAr ;
end ;

Function TFRecopiTarf.WhereSQL (Art,Cli,CatArt,CatCli: Boolean) : String ;
Var StWhere : String ;
BEGIN
StWhere:='' ;
if Art then
   if StWhere<>'' then StWhere:= StWhere+ 'OR (GF_ARTICLE <>"" AND GF_TARIFARTICLE="") '
      else StWhere:= StWhere+ ' ((GF_ARTICLE <>"" AND GF_TARIFARTICLE="") ' ;
if CatArt then
   if StWhere<>'' then StWhere:=StWhere+ 'OR (GF_ARTICLE ="" AND GF_TARIFARTICLE<>"" AND GF_TIERS="" AND GF_TARIFTIERS="") '
      else StWhere:=StWhere+ ' ((GF_ARTICLE ="" AND GF_TARIFARTICLE<>"" AND GF_TIERS="" AND GF_TARIFTIERS="") ' ;
if cli then
   if StWhere<>'' then StWhere:=StWhere+ 'OR (GF_TIERS<>"" AND GF_TARIFTIERS="" AND GF_TARIFARTICLE="") '
      else StWhere:=StWhere+ ' ((GF_TIERS<>"" AND GF_TARIFTIERS="" AND GF_TARIFARTICLE="") ';
if Cli And CatArt then
   if StWhere<>'' then StWhere:=StWhere+ 'OR (GF_TIERS<>"" AND GF_TARIFTIERS="" AND GF_TARIFARTICLE<>"") '
      else StWhere:=StWhere+ ' ((GF_TIERS<>"" AND GF_TARIFTIERS="" AND GF_TARIFARTICLE<>"") ' ;
if catcli then
   if StWhere<>'' then StWhere:=StWhere+ 'OR (GF_TIERS="" AND GF_TARIFTIERS<>"" AND GF_TARIFARTICLE="") '
      else StWhere:=StWhere+ ' ((GF_TIERS="" AND GF_TARIFTIERS<>"" AND GF_TARIFARTICLE="") ' ;
if CatCli And CatArt then
   if StWhere<>'' then StWhere:=StWhere+ 'OR (GF_TIERS="" AND GF_TARIFTIERS<>"" AND GF_TARIFARTICLE<>"") '
      else StWhere:=StWhere+ ' ((GF_TIERS="" AND GF_TARIFTIERS<>"" AND GF_TARIFARTICLE<>"") ';
Result:=StWhere ;

{SQL:='Select * from tarif where ' ;
SQL:=SQL + StWhere + ') AND GF_DEPOT="'+Depot+'"'+
         ' AND GF_TARFMODE="'+IntToStr(TarifMode)+'" AND GF_FERME="-"' ;
Result:=SQL ;}
END ;

Function TFRecopiTarf.CreerTobDepot(Depot:String) :Tob ;
Var SQL,StWhere: String;
TobDepot: TOB;
Q: TQuery ;
TarifMode: Integer ;
begin
TarifMode:=RecupCodeTarifMode(TYPETARIF.Value,PERIODE.Value,'VTE') ;
Stwhere:=WhereSQL(CB_TARFART.Checked,CB_TARFCLI.Checked,CB_TARFCATART.Checked,CB_TARFCATCLI.Checked) ;
SQL:='Select * from tarif where ' ;
SQL:=SQL + StWhere + ') AND GF_TARFMODE="'+IntToStr(TarifMode)+'" AND GF_FERME="-"' ;
SQL:=SQL + 'AND GF_DEPOT="'+Depot+'"' ;
TOBDepot:=TOB.Create('LesTarifsDepot',nil,-1) ;
Q:=OpenSQL(SQL,True) ;
if Not Q.EOF then TOBDepot.LoadDetailDB('TARIF','','',Q,False) ;
Result:=TOBDepot ;
ferme(Q) ;
end ;

Function VerifTarifExistant(TobDepot,TobTarifFille:TOB): TOB ;
var i,Compare:Integer ;
begin
result:=Nil ;
for i:=0 to TobDepot.Detail.count-1 do
  begin
  compare:=CompareTOB(TobDepot.Detail[i],TobTarifFille,'GF_ARTICLE;GF_TIERS;GF_TARIFARTICLE;GF_TARIFTIERS') ;
  if compare=0 then
    begin
    Result:=TobDepot.Detail[i] ;
    Exit ;
    end ;
  end ;
end ;

end.
