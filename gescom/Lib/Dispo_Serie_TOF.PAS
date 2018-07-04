{***********UNITE*************************************************
Auteur  ...... : Joël TRIFILIEFF
Créé le ...... : 14/03/2002
Modifié le ... : 29/03/2002
Description .. : Source TOF de la FICHE : DISPO_SERIE ()
Mots clefs ... : TOF;DISPO_SERIE
*****************************************************************}
Unit DISPO_SERIE_TOF ;

INTERFACE

Uses StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
     maineagl,emul,
{$ELSE}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Mul,FE_main,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,Vierge,UTOF,
     UTOB,AGLInit,AglInitGC,M3FP,HTB97,Grids , Windows;

function GCLanceFiche_DispoSerie(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_DISPO_SERIE = Class (TOF)
  private
    Article,Depot,Lot,Prov,Chps,NSerieDoublon : String;
    TOBSerie,TOBSerieDel                      : TOB;
    G_SER                                     : THGrid;
    QtEcart,QtProv,QtSerie,PosNvLigGrid       : integer;
    NColSerie,NColRsv,NColPrepa               : integer;
    Qry                                       : TQuery;
    BImprimer,BFerme                          : TToolbarButton97;
    procedure BImprimerOnClick(Sender: TObject);
    procedure BFermeOnClick(Sender: TObject);
    procedure CalcTotaux;
    procedure G_SEROnCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_SEROnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure EtudieColsListe ;
    procedure ModeSaisie;
    procedure ConstruitTOB;
    procedure CreeTOB(NumLig: integer);
    procedure MajTOB(NumLig : integer);
    procedure ValiderSerie;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    function TestModif : Boolean;
    function MsgEnrModif : Integer;

  public
    procedure OnArgument (StArgument : String ) ; override ;
    procedure OnNew    ; override ;
    procedure OnDelete ; override ;
    procedure OnUpdate ; override ;
    procedure OnLoad   ; override ;
    procedure OnClose  ; override ;
    procedure OnCancel ; override ;
  end ;

implementation

function GCLanceFiche_DispoSerie(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:='';
if Nat='' then exit;
if Cod='' then exit;
Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

{==============================================================================================}
{================================== Procédure de la TOF =======================================}
{==============================================================================================}
procedure TOF_DISPO_SERIE.OnArgument (StArgument : String ) ;
var Param,Extract : string;
    PosEgal       : integer;
begin
  Inherited ;
Param   := uppercase(StArgument); ReadTokenSt(Param);
while Param<>'' do
      begin
      Extract := ReadTokenSt(Param);
      PosEgal := Pos('=',Extract);
      if pos('PROV=',Extract)>0 then         Prov    := copy(Extract,PosEgal+1,length(Extract))
      else if pos('ARTICLE=',Extract)>0 then Article := copy(Extract,PosEgal+1,length(Extract))
      else if pos('DEPOT=',Extract)>0 then   Depot   := copy(Extract,PosEgal+1,length(Extract))
      else if pos('LOT=',Extract)>0 then     Lot     := copy(Extract,PosEgal+1,length(Extract))
      else if pos('QTE=',Extract)>0 then     QtProv  := StrToInt(copy(Extract,PosEgal+1,length(Extract)));
      end;
Chps      := 'GQS_IDSERIE;GQS_ENRESERVECLI;GQS_ENPREPACLI';
BFerme    := TToolBarButton97(GetControl('BFerme'));    BFerme.OnClick    := BFermeOnClick;
BImprimer := TToolBarButton97(GetControl('BImprimer')); BImprimer.OnClick := BImprimerOnClick;
BImprimer.Visible := False;
G_SER := THGrid(GetControl('G_SER'));
G_SER.ColCount    := 4;
G_SER.OnCellEnter := G_SEROnCellEnter;
G_SER.OnCellExit  := G_SEROnCellExit;
G_SER.Tag:=1;
G_SER.ListeParam  := 'GCDISPOSERIE';
TFVierge(Ecran).HMTrad.ResizeGridColumns(G_SER) ;

Qry := OpenSQL('Select GA_LIBELLE from ARTICLE where GA_ARTICLE="'+Article+'"',true,-1,'',true);
SetControlText('ARTICLE',Qry.FindField('GA_LIBELLE').AsString);
Ferme(Qry);
Qry := OpenSQL('select GDE_LIBELLE from DEPOTS where GDE_DEPOT="'+Depot+'"',true,-1,'',true);
SetControlText('DEPOT',Qry.FindField('GDE_LIBELLE').AsString);
Ferme(Qry);
SetControlText('TOT_ART',IntToStr(QtProv));

If Prov='LOT' then
   begin
   SetControlText('LOT',Lot);
   SetControlText('QTEART','Qté du lot');
   end
else if prov='DISPO' then
   begin
   SetControlVisible('LOT',false);
   SetControlVisible('TGQS_LOT',false);
   SetControlText('QTEART','Qté de l''article');
   end;

TOBSerie    := TOB.Create('',nil,-1);
TOBSerieDel := TOB.Create('',nil,-1);
TForm(Ecran).OnKeyDown := FormKeyDown;
ConstruitTOB;
EtudieColsListe;
end ;

procedure TOF_DISPO_SERIE.OnNew ;
var NumLigTob : integer;
begin
  Inherited ;
MajTOB(PosNvLigGrid); //Maj de la TOB pour la ligne précédente

if QtEcart=0 then
   PGIBox('La quantité maximum est atteinte.','Ajout impossible')
else begin
   if TOBSerie.Detail.count>0 then //La TOB n'est pas vide, insertion d'une nouvelle ligne
      begin
      PosNvLigGrid := G_SER.Row;
      NumLigTob    := PosNvLigGrid-1;
      G_SER.InsertRow(PosNvLigGrid);
      end
   else
      begin
      PosNvLigGrid  := 1;
      NumLigTob := 0;
      end;
   CreeTOB(NumligTob);
   CalcTotaux;
   end;
end;

procedure TOF_DISPO_SERIE.OnCancel ;
var TobRech : TOB;
begin
  Inherited ;
TobRech := TobSerie.FindFirst(['Nouvelle'],['X'],True);
if (TobRech<>nil) or (TobSerieDel.Detail.count>0) then
   begin
   if PGIAsk('Voulez-vous annuler les modifications ?','Confirmation')=mrYes then
      begin
      G_SER.VidePile(False);
      TOBSerie.ClearDetail;
      TOBSerieDel.ClearDetail;
      ConstruitTOB;
      end;
   end;
end;

procedure TOF_DISPO_SERIE.OnDelete ;
var Lig     : integer;
    TOBSDel : TOB;
begin
  Inherited ;
if TOBSerie.Detail.count=0 then exit;

Lig := G_SER.Row;
if G_SER.CellValues[NColRsv,Lig]='X' then
   PGIBox('Impossible de supprimer une ligne en "Réservé client".','Erreur')
else if G_SER.CellValues[NColPrepa,Lig]='X' then
   PGIBox('Impossible de supprimer une ligne en "Préparation client".','Erreur')
else
    begin
    if G_SER.CellValues[NColSerie,Lig]<>'' then     //Si un n° est saisi, demande confirmation
       if PGIAsk('Voulez-vous supprimer cette ligne ?','Confirmation')=mrNo then exit;

    if TOB(G_SER.Objects[0,Lig]).GetValue('Nouvelle')='-' then //Copy dans la TOB delete si ce n'est pas une nouvelle ligne
       begin
       TOBSDel := TOB.Create('',TOBSerieDel,-1);
       TOBSDel.Dupliquer(TOB(G_SER.Objects[0,Lig]),False,True,True);
       end;

    TOB(G_SER.Objects[0,Lig]).Free;   // Destruction de la ligne de la TOB
    G_SER.DeleteRow(Lig);             //Destruction de la ligne du Grid

   If TOBSerie.Detail.count=0 then   //Si la TOB est vide (donc plus de ligne dans le Grid) ajout d'une ligne vide dans le Grid
       begin
       G_SER.RowCount  := 2;
       G_SER.FixedRows := 1;
       G_SER.Row       := 1;
       end;

    CalcTotaux;
    ModeSaisie;
    end;
end ;

procedure TOF_DISPO_SERIE.OnUpdate ;
var io    : TIOErr ;
    Ligne : integer;
begin
  Inherited ;
MajTOB(G_SER.Row);
io := Transactions(ValiderSerie,0);
case io of
     oeOk      : begin
                 for Ligne := 0 to TOBSerie.Detail.count-1 do
                     if (TOBSerie.Detail[Ligne].Getvalue('GQS_IDSERIE')<>'') and (TOBSerie.Detail[Ligne].Getvalue('Nouvelle')='X') then
                        TOBSerie.Detail[Ligne].PutValue('Nouvelle','-');
                 end;
     oeUnknown : begin
                 PGIBox('Le numéro de série "'+NSerieDoublon+'" est en doublon pour cet article.','Erreur');
                 TForm(Ecran).ModalResult:=0;
                 LastError := -1;
                 exit;
                 end;
     end;
end ;

procedure TOF_DISPO_SERIE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_DISPO_SERIE.Onclose ;
begin
  Inherited ;
end ;

procedure TOF_DISPO_SERIE.ConstruitTOB;
Var Cpt : integer;
begin
for Cpt:=2 to 3 do
    begin
    G_SER.ColLengths[Cpt] := -1;                         //Non saisissable
    G_SER.ColTypes[Cpt]   := 'B';                        //B = Boolean
    G_SER.ColFormats[Cpt] := IntToStr(integer(csCoche)); //Affichage mouette ou rien
    end;

if Prov='DISPO' then
   Qry := OpenSQL('select * from DISPOSERIE where GQS_ARTICLE="'+Article+'" and GQS_DEPOT="'+Depot+'" order by GQS_IDSERIE',True,-1,'',true)
else if Prov='LOT'then
   Qry := OpenSQL('select * from DISPOSERIE where GQS_ARTICLE="'+Article+'" and GQS_DEPOT="'+Depot+'" and GQS_NUMEROLOT="'+Lot+'" order by GQS_IDSERIE',True,-1,'',true);
TOBSerie.LoadDetailDB('DISPOSERIE','','',Qry,False,False);
Ferme(Qry);
for cpt := 0 to TOBSerie.Detail.count-1 do  //Ajout champs supplémentaire (- =déjà existante, X = nouvelle ligne)
   begin
   TOBSerie.Detail[Cpt].AddChampSup('Nouvelle',False);
   TOBSerie.Detail[Cpt].PutValue('Nouvelle','-');
   end;

CalcTotaux;

if QtSerie>0 then
   begin
   G_SER.RowCount := QtSerie+1;
   TOBSerie.PutGridDetail(G_SER,True,True,Chps,True);
   ModeSaisie;
   end;
end;

procedure TOF_DISPO_SERIE.MajTOB(NumLig : integer);
begin
if (NumLig>0)
   and (TOBSerie.Detail.count>0)
   and (TOB(G_SER.Objects[0,NumLig]).GetValue('Nouvelle')='X') then
   TOB(G_SER.Objects[0,NumLig]).PutValue('GQS_IDSERIE',uppercase(G_SER.Cellvalues[NColSerie,NumLig]));
end;

procedure TOF_DISPO_SERIE.CreeTOB(NumLig: integer);
var TOBl : TOB;
begin
TOBl := TOB.Create('DISPOSERIE',TOBSerie,NumLig);
TOBl.PutValue('GQS_ARTICLE',Article);
TOBl.PutValue('GQS_DEPOT',Depot);
TOBl.PutValue('GQS_NUMEROLOT',Lot);
TOBl.PutValue('GQS_IDSERIE','');
TOBl.PutValue('GQS_ENRESERVECLI','-');
TOBl.PutValue('GQS_ENPREPACLI','-');
TOBSerie.Detail[NumLig].AddChampSup('Nouvelle',False);
TOBSerie.Detail[NumLig].PutValue('Nouvelle','X');
TOBl.PutLigneGrid(G_SER,PosNvLigGrid,False,False,Chps);

G_SER.Row := PosNvLigGrid;
ModeSaisie;
end;

procedure TOF_DISPO_SERIE.CalcTotaux;
begin
QtSerie := TOBSerie.Detail.count;
QtEcart := QtSerie-QtProv;
SetControlText('TOT_SERIE',IntToStr(QtSerie));
SetControlText('TOT_ECART',IntToStr(QtEcart));
end;

procedure TOF_DISPO_SERIE.ModeSaisie;
Var i_ind : integer ;
begin
if TOBSerie.Detail.count>0 then
   begin
   i_ind:=G_SER.row-1;
   if TOBSerie.Detail[i_ind].GetValue('Nouvelle')='X' then
      begin
      G_SER.Options:=G_SER.Options+[goEditing];
      G_SER.ShowEditor;
      end else
      G_SER.Options:=G_SER.Options-[goEditing];
  end;
end;

function TOF_DISPO_SERIE.TestModif : Boolean;
var Cpt : integer;
begin
Result:=False;
if TOBSerieDel.Detail.count>0 then
   Result := True
else
   begin
   if TOBSerie.IsOneModifie then
      begin
      for Cpt := 0 to TOBSerie.Detail.count-1 do
          begin
          if (TOBSerie.Detail[Cpt].GetValue('Nouvelle')='X') and (G_SER.CellValues[NColSerie,Cpt+1]<>'') then
             begin
             Result := True;
             Break;
             end;
          end;
      end;
   end;
end;

function TOF_DISPO_SERIE.MsgEnrModif : Integer;
var Msg : string ;
begin
Msg    := '0;?CAPTION?;Voulez-vous enregistrer les modifications ?;Q;YNC;Y;C;';
Result := HShowMessage(Msg, Ecran.Caption, '');
end;

{==============================================================================================}
{==================================== Gestion du Grid =========================================}
{==============================================================================================}
procedure TOF_DISPO_SERIE.EtudieColsListe ;
Var NomCol,LesCols : String ;
    icol,ichamp    : integer ;
begin
LesCols := G_SER.Titres[0];
icol:=1 ;
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
        begin
        ichamp:=ChampToNum(NomCol) ;
        if ichamp>=0 then
           begin
           if NomCol='GQS_IDSERIE' then NColSerie := icol
           else if NomCol='GQS_ENRESERVECLI' then NColRsv := icol
           else if NomCol='GQS_ENPREPACLI' then NColPrepa := icol
           end;
        end ;
    Inc(icol) ;
    Until ((LesCols='') or (NomCol='')) ;
end ;

procedure TOF_DISPO_SERIE.G_SEROnCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
ModeSaisie;
end;

procedure TOF_DISPO_SERIE.G_SEROnCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
MajTOB(Arow);
end;

{==============================================================================================}
{================================ Evenements des boutons ======================================}
{==============================================================================================}
procedure TOF_DISPO_SERIE.BFermeOnClick(Sender: TObject);
var Reponse : integer;
begin
MajTOB(G_SER.Row);
Reponse := MsgEnrModif;
if TestModif then
   Case Reponse of
        mrYes    : begin
                   LastError := 0;
                   OnUpdate;
                   if LastError<>0 then
                      exit
                   else
                      begin
                      TOBSerieDel.Free; TOBSerie.Free;
                      TForm(Ecran).ModalResult := 2;
                      end;
                   end;
        mrNo     : begin
                   TOBSerieDel.Free; TOBSerie.Free;
                   TForm(Ecran).ModalResult := 2;
                   end;
        mrCancel : begin
                   TForm(Ecran).ModalResult := 0;
                   exit;
                   end;
        end;
end ;

procedure TOF_DISPO_SERIE.BImprimerOnClick(Sender: TObject);
begin
end;

{==============================================================================================}
{======================================= Validation ===========================================}
{==============================================================================================}
procedure TOF_DISPO_SERIE.ValiderSerie;
var Ligne  : integer;
    NSerie : string;
begin
if TOBSerieDel.Detail.count>0 then
   begin
   TOBserieDel.DeleteDB(False);
   TOBSerieDel.ClearDetail;
   end;

for Ligne := 0 to TOBSerie.Detail.count-1 do
   begin
   if (TOBSerie.Detail[Ligne].Getvalue('GQS_IDSERIE')<>'') and (TOBSerie.Detail[Ligne].Getvalue('Nouvelle')='X') then
      begin
      NSerie := TOBSerie.Detail[Ligne].Getvalue('GQS_IDSERIE');
      if not ExisteSQL('Select GQS_IDSERIE from DISPOSERIE Where GQS_ARTICLE="'+Article+'" and GQS_IDSERIE="'+NSerie+'"') then  //Test si existe déjà
         TOBSerie.Detail[Ligne].InsertDB(nil,False)
      else
         begin
         NSerieDoublon := NSerie;
         V_PGI.IoError := oeUnknown;
         Exit;
         end;
      end;
   end;

end;

{==============================================================================================}
{================================== Touches de contrôle =======================================}
{==============================================================================================}
procedure TOF_DISPO_SERIE.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
Case Key of
   VK_INSERT : OnNew;
   VK_DELETE : begin
               if (Shift = [ssCtrl]) then
                  begin
                  Key := 0;
                  OnDelete;
                  end;
               end;
   end;
end;

Initialization
  registerclasses ( [ TOF_DISPO_SERIE ] ) ;
end.


