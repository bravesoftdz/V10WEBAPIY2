unit CalculStock;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, HSysMenu, HTB97, StdCtrls, Hctrls, Mask, hmsgbox, UTOB, DBTables,
  ed_tools, Hent1, HPanel, UIUTil;

procedure AppelCalculStock();

type
  TFCalculStock = class(TFVierge)
    TChkStockNull: TCheckBox;
    TChkDepotLigne: TCheckBox;
    TChkPhysique: TCheckBox;
    TChkVenteFFO: TCheckBox;
    TChkEntreesSorties: TCheckBox;
    TChkEcartINV: TCheckBox;
    TChkTransfert: TCheckBox;
    TChkLivreFou: TCheckBox;
    TChkLivreClient: TCheckBox;
    TChkPrepaCli: TCheckBox;
    TChkReserveFou: TCheckBox;
    TchkReserveCli: TCheckBox;
    TGP_DEPOT: THLabel;
    TChkClotureStock: TCheckBox;
    DtVenteFFO: THCritMaskEdit;
    DtEntreesSorties: THCritMaskEdit;
    DtEcartINV: THCritMaskEdit;
    DtTransfert: THCritMaskEdit;
    DtLivreFou: THCritMaskEdit;
    DtLivreClient: THCritMaskEdit;
    DtPrepaCli: THCritMaskEdit;
    DtReserveFou: THCritMaskEdit;
    DtReserveCli: THCritMaskEdit;
    TDateDepart: THLabel;
    GP_DEPOT: THMultiValComboBox;
    TChkClotureDate: TCheckBox;
    DtClotureDate: THCritMaskEdit;
    TChkEgal: TCheckBox;
    TChkEgalDate: TCheckBox;
    TChkTrace: TCheckBox;
    TOk: TLabel;
    DtLastDateCloture: THCritMaskEdit;
    TChkVenteFFO_0: TCheckBox;
    TChkEntreesSorties_0: TCheckBox;
    TChkEcartINV_0: TCheckBox;
    TChkTransfert_0: TCheckBox;
    TChkLivreFou_0: TCheckBox;
    TChkLivreClient_0: TCheckBox;
    TChkPrepaCli_0: TCheckBox;
    TChkReserveFou_0: TCheckBox;
    TchkReserveCli_0: TCheckBox;
    TChkPhysique_0: TCheckBox;
    HLabel1: THLabel;
    procedure BValiderClick(Sender: TObject);
    procedure TChkPhysiqueClick(Sender: TObject);
    procedure TChkClotureStockClick(Sender: TObject);
  private
    { Déclarations privées }
    BDateCloture,BVerifStock,BVerifDepot,BPhysique,BFFO,BEEX,BInv,
    BTransfert,BLivFou,BLivCli,BClotureDate : Boolean;
    SupEgal1,SupEgal2 : string;
    procedure LanceCalcul;
    function  ChargeListeDepot : string;
    function  NbZonesCochees : integer;
    procedure BloqueEtSelectionZonesStock(BPhys,Force:boolean);
    procedure BloqueEtSelectionZonesAZero(BPhys:boolean);
    function  FindDateCloture(TT : TQRProgressForm; LeDepot : string): TDateTime;
    procedure VerifDepot(TT : TQRProgressForm; LeDepot : string);
    procedure VerifStock(TT : TQRProgressForm; LeDepot : string);
    procedure InitialiseDispo(TT:TQRProgressForm; Depot:string);
    procedure InitialiseDispo2(TT:TQRProgressForm; Depot:string; DateClot:TDateTime);
    function  SQLZone(Zone:string; DateClot:TDateTime): string;
    procedure CalculZoneFFO(TT:TQRProgressForm; Depot:string; DateClot:TDateTime);
    procedure CalculZoneEEX(TT:TQRProgressForm; Depot:string; DateClot:TDateTime);
    procedure CalculZoneInv(TT:TQRProgressForm; Depot:string; DateClot:TDateTime);
    procedure CalculZoneTransfert(TT:TQRProgressForm; Depot:string; DateClot:TDateTime);
    procedure CalculZoneLivFou(TT:TQRProgressForm; Depot:string; DateClot:TDateTime);
    procedure CalculZoneLivCli(TT:TQRProgressForm; Depot:string; DateClot:TDateTime);
    procedure CalculZonePhysique(TT:TQRProgressForm; Depot:string);
    procedure ClotureStockDate(TT:TQRProgressForm; Depot:string);
    function  RequeteAvecDateCloture(Depot,NomChamp,Nature:string; TheDate:TDateTime; Negatif:Boolean=False): String;
    function  RequeteAvecDate(Depot,NomChamp,Nature:string; TheDate:TDateTime; Negatif:Boolean=False): String;
    function  RequeteSansDateCloture(Depot,NomChamp,Nature:string; Negatif:Boolean=False): String;
    function  RequeteAvecDate2(Depot,NomChamp,Nature:string; TheDate:TDateTime): String;
    function  RequeteSansDateCloture2(Depot,NomChamp,Nature:string): String;
    function  ExistePieceNature(Nature,Etablissement:string; LaDate:TDateTime): boolean;
    procedure TraceLeRecalcul(Etape:integer; First:Boolean; Info:string='');
  public
    { Déclarations publiques }
  end;

var
  FCalculStock: TFCalculStock;

implementation

{$R *.DFM}

procedure AppelCalculStock();
var X : TFCalculStock ;
    PP  : THPanel ;
begin
SourisSablier;
PP:=FindInsidePanel;
X :=TFCalculStock.Create(Application);
if PP=Nil then
   begin
    try
      X.ShowModal ;
    finally
      X.Free ;
    end ;
   SourisNormale ;
   end
else
   begin
   //X.Autosize:=True;
   InitInside(X,PP); X.Show;
   end;
end;

procedure TFCalculStock.TraceLeRecalcul(Etape:integer; First:Boolean; Info:string='');
var FichierASCII : TextFile;
    JH,Emplacement,St : string;
begin
  Emplacement := ExtractFilePath(Application.EXEName) + 'LogRecalculStock.log';
  AssignFile(FichierASCII, Emplacement);
  JH := FormatDateTime('c',Now);
  case Etape of
    1 : begin
        if FileExists(Emplacement) then Append(FichierASCII) else Rewrite(FichierASCII);
        WriteLn(FichierASCII,JH+' : Début traitement');
        end;
    2 : begin
        Append(FichierASCII); WriteLn(FichierASCII,JH+' : Boutique traitée : '+Info);
        end;
    3 : begin
        if First then St:='Recherche date clotûre' else St:='Ok : Date clôture trouvée '+Info;
        Append(FichierASCII); WriteLn(FichierASCII,JH+' : '+St);
        end;
    4 : begin
        if First then St:='Vérif dépôt' else St:='Ok'+' : '+Info+' enregistrements traités';
        Append(FichierASCII); WriteLn(FichierASCII,JH+' : '+St);
        end;
    5 : begin
        if First then St:='Vérif stock' else St:='Ok'+' : '+Info+' enregistrements traités';
        Append(FichierASCII); WriteLn(FichierASCII,JH+' : '+St);
        end;
    6 : begin
        if First then St:='Initialise dispo' else St:='Ok'+' : '+Info+' enregistrements traités';
        Append(FichierASCII); WriteLn(FichierASCII,JH+' : '+St);
        end;
    7 : begin
        if First then St:='Recalcul zone FFO' else St:='Ok'+' : '+Info+' enregistrements traités';
        Append(FichierASCII); WriteLn(FichierASCII,JH+' : '+St);
        end;
    8 : begin
        if First then St:='Recalcul zone EEX' else St:='Ok'+' : '+Info+' enregistrements traités';
        Append(FichierASCII); WriteLn(FichierASCII,JH+' : '+St);
        end;
    9 : begin
        if First then St:='Recalcul zone INV' else St:='Ok'+' : '+Info+' enregistrements traités';
        Append(FichierASCII); WriteLn(FichierASCII,JH+' : '+St);
        end;
    10: begin
        if First then St:='Recalcul zone Transfert' else St:='Ok'+' : '+Info+' enregistrements traités';
        Append(FichierASCII); WriteLn(FichierASCII,JH+' : '+St);
        end;
    11: begin
        if First then St:='Recalcul zone Liv Fou' else St:='Ok'+' : '+Info+' enregistrements traités';
        Append(FichierASCII); WriteLn(FichierASCII,JH+' : '+St);
        end;
    12: begin
        if First then St:='Recalcul zone Liv Cli' else St:='Ok'+' : '+Info+' enregistrements traités';
        Append(FichierASCII); WriteLn(FichierASCII,JH+' : '+St);
        end;
    13: begin
        if First then St:='Recalcul zone GQ_PHYSIQUE' else St:='Ok'+' : '+Info+' enregistrements traités';
        Append(FichierASCII); WriteLn(FichierASCII,JH+' : '+St);
        end;
    14: begin
        if First then St:='Clôture stock à date' else St:='Ok'+' : '+Info+' enregistrements traités';
        Append(FichierASCII); WriteLn(FichierASCII,JH+' : '+St);
        end;
    15: begin
        Append(FichierASCII); WriteLn(FichierASCII,JH+' : Fin traitement');
        end;
    end;
  CloseFile(FichierASCII);
end;

procedure TFCalculStock.BValiderClick(Sender: TObject);
begin
  inherited;
  if PGIAsk('Etes-vous sûre de vouloir mettre à jour vos stocks ? Cette opération est irréversible.',FCalculStock.Caption)
     = MrYes then LanceCalcul;
end;

function TFCalculStock.ChargeListeDepot : string;
var i : integer;
    Liste : String;
begin
  // Recherche de la liste des dépôts sélectionnés
  Liste := '' ;
  if GP_DEPOT.Text = TraduireMemoire ('<<Tous>>') then
    begin
    For i:=0 to GP_DEPOT.Items.Count-1 do
      begin
      if Liste<>'' then Liste := Liste+';';
      Liste := Liste+GP_DEPOT.Values[i];
      end;
    if Liste<>'' then Liste := Liste+';';
    end
  else Liste := GP_DEPOT.Text;
  Result:=Liste;
end;

function TFCalculStock.NbZonesCochees : integer;
var i : integer;
begin
  i:=0;
  BDateCloture:=TChkClotureStock.Checked; if BDateCloture then Inc(i);
  BVerifStock:=TChkStockNull.Checked; if BVerifStock then Inc(i);
  BVerifDepot:=TChkDepotLigne.Checked; if BVerifDepot then Inc(i);
  BPhysique:=TChkPhysique.Checked; if BPhysique then Inc(i);
  BFFO:=TChkVenteFFO.Checked; if BFFO then Inc(i);
  BEEX:=TChkEntreesSorties.Checked; if BEEX then Inc(i);
  BInv:=TChkEcartINV.Checked; if BInv then Inc(i);
  BTransfert:=TChkTransfert.Checked; if BTransfert then Inc(i);
  BLivFou:=TChkLivreFou.Checked; if BLivFou then Inc(i);
  BLivCli:=TChkLivreClient.Checked; if BLivCli then Inc(i);
  {BPrepaLiv:=TChkPrepaCli.Checked; if BPrepaLiv then Inc(i);
  BCommFou:=TChkReserveFou.Checked; if BCommFou then Inc(i);
  BCommCli:=TchkReserveCli.Checked; if BCommCli then Inc(i);}
  BClotureDate:=TChkClotureDate.Checked; if BClotureDate then Inc(i);
  Result := i;
end;

function TFCalculStock.FindDateCloture(TT : TQRProgressForm; LeDepot : string): TDateTime;
var Q : TQuery;
begin
  TraceLeRecalcul(3,True);
  TT.SubText := 'Recherche la dernière date de clôture'; TT.Value:=TT.Value+1;
  if StrToDate(DtLastDateCloture.Text)=iDate1900 then
    begin
    Q:=OpenSQL('SELECT MAX(GQ_DATECLOTURE) AS MaxDate FROM DISPO WHERE GQ_CLOTURE="X" AND GQ_DEPOT="'+LeDepot+'"',True);
    if Not Q.Eof then
      begin
      if Q.FindField('MaxDate').AsString='' then Result:=iDate1900
      else Result:=Q.FindField('MaxDate').AsDateTime;
      end
    else Result:=iDate1900;
    Ferme(Q);
    //DtLastDateCloture.Text:=DateToStr(UsDateTimeToDateTime(Result));
    end
  else Result:=StrToDate(DtLastDateCloture.Text);
  TraceLeRecalcul(3,False,DateToStr(Result));
end;

procedure TFCalculStock.VerifDepot(TT : TQRProgressForm; LeDepot : string);
var SQL : String;
    Nb : Integer;
begin
  TraceLeRecalcul(4,True);
  TT.SubText := 'Vérifie les dépôts sur tous les documents'; TT.Value:=TT.Value+1;
  SQL := 'Update ligne set gl_depot=(select gp_depotdest from piece where '+
         'gl_naturepieceg=gp_naturepieceg and gl_souche=gp_souche '+
         'and gl_numero=gp_numero and gl_indiceg=gp_indiceg) '+
         ', gl_etablissement=(select gp_depotdest from piece where '+
         'gl_naturepieceg=gp_naturepieceg and gl_souche=gp_souche '+
         'and gl_numero=gp_numero and gl_indiceg=gp_indiceg) '+
         'where gl_naturepieceg="TRE"';
  Nb := ExecuteSQL(SQL);
  TraceLeRecalcul(4,False,IntToStr(Nb));
end;

procedure TFCalculStock.VerifStock(TT : TQRProgressForm; LeDepot : string);
var SQL,Info : String;
    Nb : integer;
begin
  TraceLeRecalcul(5,True);
  TT.SubText := 'Vérifie l''existence de toutes les fiches stocks'; TT.Value:=TT.Value+1;
  //Nettoyage de la table dispo
  SQL := 'Delete from dispo where gq_depot is null';
  Nb := ExecuteSQL(SQL); Info := IntToStr(Nb);
  SQL := 'Delete from dispo where not exists (Select ga_article from article where ga_article=gq_article)';
  Nb := ExecuteSQL(SQL); Info := Info+' + '+IntToStr(Nb);

  //On ajoute un enregistrement dans la table dispo pour tous les articles ayant un
  //mouvement et n'ayant pas de fiche stock.
  //On laisse les natures de pieces comme ALF qui n'impacte pas le stock pour un traitement plus rapide
  SQL := 'INSERT INTO DISPO '+
         'select DISTINCT GL_ARTICLE, GL_DEPOT, 0.0,0.0,"01/01/1900",0.0,0.0,0.0,0.0,0.0,'+
         '0.0,0.0,0.0,0.0,0.0,GA_DPA,GA_PMAP,GA_DPR,GA_PMRP,'+
         '"'+UsDateTime(Date)+'","'+UsDateTime(Date)+
         '","'+UsDateTime(Date)+'","CEG",0.0,'+
         '"-","01/01/1900","",0.0,0.0,0.0,0.0,0.0 '+
         'from LIGNE left join ARTICLE on GA_ARTICLE=GL_ARTICLE '+
         'where GL_TYPEARTICLE="MAR" and GL_TYPEDIM<>"GEN" and '+
         'NOT Exists (select GQ_ARTICLE from DISPO '+
         'where GQ_ARTICLE=GL_ARTICLE and GQ_DEPOT=GL_DEPOT and GQ_CLOTURE="-")';
  Nb := ExecuteSQL(SQL); Info := Info+' + '+IntToStr(Nb);
  TraceLeRecalcul(5,False,Info);
end;

function TFCalculStock.SQLZone(Zone:string; DateClot:TDateTime): string;
begin
  Result := Zone+'=(select '+Zone+' from dispo as d2 '+
         'where d2.gq_article=dispo.gq_article and d2.gq_depot=dispo.gq_depot '+
         'and d2.gq_cloture="X" and d2.gq_datecloture="'+UsDateTime(DateClot)+'")';
end;

procedure TFCalculStock.InitialiseDispo(TT:TQRProgressForm; Depot:string);
var SQL,St,Zone : String;
begin
  TraceLeRecalcul(6,True);
  TT.SubText := 'Initialise les stocks';
  Zone:='';
  if TChkVenteFFO_0.Checked then Zone:='gq_venteffo=0';
  if TChkEntreesSorties_0.Checked then
    if Zone='' then Zone:='gq_entreesorties=0' else Zone:=Zone+',gq_entreesorties=0';
  if TChkEcartINV_0.Checked then
    if Zone='' then Zone:='gq_ecartinv=0' else Zone:=Zone+',gq_ecartinv=0';
  if TChkTransfert_0.Checked then
    if Zone='' then Zone:='gq_transfert=0' else Zone:=Zone+',gq_transfert=0';
  if TChkLivreFou_0.Checked then
    if Zone='' then Zone:='gq_livrefou=0' else Zone:=Zone+',gq_livrefou=0';
  if TChkLivreClient_0.Checked then
    if Zone='' then Zone:='gq_livreclient=0' else Zone:=Zone+',gq_livreclient=0';
  if TChkPhysique_0.Checked then
    if Zone='' then Zone:='gq_physique=0' else Zone:=Zone+',gq_physique=0';

  if Zone='' then exit;
  SQL := 'Update dispo set '+Zone+' '+
         'where gq_depot="'+Depot+'" and gq_cloture="-"';
  St:=St+'+'+IntToStr(ExecuteSQL(SQL));
  TraceLeRecalcul(6,False,St);
end;

procedure TFCalculStock.InitialiseDispo2(TT:TQRProgressForm; Depot:string; DateClot:TDateTime);
var SQL,St : String;
begin
  TraceLeRecalcul(6,True);
  TT.SubText := 'Initialise les stocks';
  SQL := 'Update dispo set gq_physique=0 ,'+
  SQLZone('gq_livreclient',DateClot)+','+SQLZone('gq_livrefou',DateClot)+','+
  SQLZone('gq_transfert',DateClot)+','+SQLZone('gq_venteffo',DateClot)+','+
  SQLZone('gq_entreesorties',DateClot)+','+SQLZone('gq_ecartinv',DateClot)+
  ' where gq_depot="'+Depot+'" and gq_cloture="-"';
  St:=IntToStr(ExecuteSQL(SQL));

  SQL := 'Update dispo set gq_physique=0,'+
         //gq_reservecli=0,gq_reservefou=0,gq_prepacli=0,'+
         'gq_livreclient=0,gq_livrefou=0,gq_transfert=0,gq_venteffo=0,'+
         'gq_entreesorties=0,gq_ecartinv=0 '+
         'where gq_depot="'+Depot+'" and gq_cloture="-" and gq_livreclient IS NULL';
  St:=St+'+'+IntToStr(ExecuteSQL(SQL));
  TraceLeRecalcul(6,False,St);
end;


function TFCalculStock.RequeteAvecDateCloture(Depot,NomChamp,Nature:string; TheDate:TDateTime; Negatif:Boolean=False) : String;
var SQL,ZoneNegative : string;
begin
  if Negatif then ZoneNegative:='* (-1)' else ZoneNegative:='';
  SQL := 'UPDATE DISPO SET '+NomChamp+'='+NomChamp+' + '+
  '(select sum(gl_qtestock)'+ZoneNegative+' from ligne where '+
  'ligne.gl_article=dispo.gq_article and gl_naturepieceg="'+Nature+'" and '+
  'gl_datepiece'+SupEgal1+'"'+UsDateTime(TheDate)+'" and ligne.gl_depot=dispo.gq_depot) '+
  'where gq_depot="'+Depot+'" and gq_cloture="-" and NOT '+
  '(select sum(gl_qtestock)'+ZoneNegative+' from ligne where '+
  'ligne.gl_article=dispo.gq_article and gl_naturepieceg="'+Nature+'" and '+
  'gl_datepiece'+SupEgal1+'"'+UsDateTime(TheDate)+'" and ligne.gl_depot=dispo.gq_depot) '+
  'IS NULL';
  Result := IntToStr(ExecuteSQL(SQL));
end;

function TFCalculStock.RequeteAvecDate(Depot,NomChamp,Nature:string; TheDate:TDateTime; Negatif:Boolean=False): String;
var SQL,ZoneNegative : string;
begin
  if Negatif then ZoneNegative:='* (-1)' else ZoneNegative:='';
  SQL := 'UPDATE DISPO SET '+NomChamp+'='+
  '(select sum(gl_qtestock)'+ZoneNegative+' from ligne where '+
  'ligne.gl_article=dispo.gq_article and gl_naturepieceg="'+Nature+'" and '+
  'gl_datepiece'+SupEgal2+'"'+UsDateTime(TheDate)+'" and ligne.gl_depot=dispo.gq_depot) '+
  'where gq_depot="'+Depot+'" and gq_cloture="-" and NOT '+
  '(select sum(gl_qtestock)'+ZoneNegative+' from ligne where '+
  'ligne.gl_article=dispo.gq_article and gl_naturepieceg="'+Nature+'" and '+
  'gl_datepiece'+SupEgal2+'"'+UsDateTime(TheDate)+'" and ligne.gl_depot=dispo.gq_depot) '+
  'IS NULL';
  Result := IntToStr(ExecuteSQL(SQL));
end;

function TFCalculStock.RequeteSansDateCloture(Depot,NomChamp,Nature:string; Negatif:Boolean=False): String;
var SQL,ZoneNegative : string;
begin
  if Negatif then ZoneNegative:='* (-1)' else ZoneNegative:='';
  SQL := 'UPDATE DISPO SET '+NomChamp+'='+
  '(select sum(gl_qtestock)'+ZoneNegative+' from ligne where '+
  'ligne.gl_article=dispo.gq_article and gl_naturepieceg="'+Nature+'" and '+
  'ligne.gl_depot=dispo.gq_depot) where gq_depot="'+Depot+
  '" and gq_cloture="-" and NOT '+
  '(select sum(gl_qtestock)'+ZoneNegative+' from ligne where '+
  'ligne.gl_article=dispo.gq_article and gl_naturepieceg="'+Nature+'" and '+
  'ligne.gl_depot=dispo.gq_depot) IS NULL';
  Result := IntToStr(ExecuteSQL(SQL));
end;

function TFCalculStock.RequeteAvecDate2(Depot,NomChamp,Nature:string; TheDate:TDateTime): String;
var SQL : string;
begin
  SQL := 'update dispo set '+NomChamp+'='+NomChamp+'+ '+
  '(select sum(gl_qtestock) from ligne where '+
  'ligne.gl_article=dispo.gq_article and gl_naturepieceg="'+Nature+'" and '+
  'gl_datepiece'+SupEgal2+'"'+UsDateTime(TheDate)+'" ligne.gl_depot=dispo.gq_depot) '+
  'where gq_depot="'+Depot+'" and gq_cloture="-" and NOT '+
  '(select sum(gl_qtestock) from ligne where '+
  'ligne.gl_article=dispo.gq_article and gl_naturepieceg="'+Nature+'" and '+
  'gl_datepiece'+SupEgal2+'"'+UsDateTime(TheDate)+'" ligne.gl_depot=dispo.gq_depot) '+
  'IS NULL';
  Result := IntToStr(ExecuteSQL(SQL));
end;

function TFCalculStock.RequeteSansDateCloture2(Depot,NomChamp,Nature:string): String;
var SQL : string;
begin
  SQL := 'update dispo set '+NomChamp+'='+NomChamp+'+'+
  '(select sum(gl_qtestock) from ligne where '+
  'ligne.gl_article=dispo.gq_article and gl_naturepieceg="'+Nature+'" and '+
  'ligne.gl_depot=dispo.gq_depot) where gq_depot="'+Depot+
  '" and gq_cloture="-" and NOT '+
  '(select sum(gl_qtestock) from ligne where '+
  'ligne.gl_article=dispo.gq_article and gl_naturepieceg="'+Nature+'" and '+
  'ligne.gl_depot=dispo.gq_depot) IS NULL';
  Result := IntToStr(ExecuteSQL(SQL));
end;

function TFCalculStock.ExistePieceNature(Nature,Etablissement:string; LaDate:TDateTime): boolean;
begin
  Result:= ExisteSQL('Select GP_NATUREPIECEG from PIECE '+
  'where GP_NATUREPIECEG="'+Nature+'" and GP_DATEPIECE>="'+UsDateTime(LaDate)+
  '" and GP_ETABLISSEMENT="'+Etablissement+'"');
end;

procedure TFCalculStock.CalculZoneFFO(TT:TQRProgressForm; Depot:string; DateClot:TDateTime);
var DD : TDateTime;
    St : string;
begin
  TraceLeRecalcul(7,True);
  TT.SubText := 'Calcul la zone ventes FFO'; TT.Value:=TT.Value+1;
  DD:=StrToDate(DtVenteFFO.Text) ;
  if DateClot<>iDate1900 then //Prise en compte du stock à la date de cloture + tous les mouvements depuis cette date
    St:=RequeteAvecDateCloture(Depot,'gq_venteffo','FFO',DateClot)
  else if DD<>iDate1900 then
    St:=RequeteAvecDate(Depot,'gq_venteffo','FFO',DD)
  else
    St:=RequeteSansDateCloture(Depot,'gq_venteffo','FFO');
  TraceLeRecalcul(7,False,St);
end;

procedure TFCalculStock.CalculZoneEEX(TT:TQRProgressForm; Depot:string; DateClot:TDateTime);
var DD : TDateTime;
    St : string;
begin
  TraceLeRecalcul(8,True);
  TT.SubText := 'Calcul la zone entrée/sorties de stock'; TT.Value:=TT.Value+1;
  DD:=StrToDate(DtVenteFFO.Text) ;
  if DateClot<>iDate1900 then DD:=DateClot;

  if ExistePieceNature('SEX',Depot,DD) then
    begin
    if DateClot<>iDate1900 then
      St:=RequeteAvecDateCloture(Depot,'gq_entreesorties','SEX',DateClot,True)
    else if DD<>iDate1900 then
      St:=RequeteAvecDate(Depot,'gq_entreesorties','SEX',DD,True)
    else
      St:=RequeteSansDateCloture(Depot,'gq_entreesorties','SEX',True);
    end
  else St := TraduireMemoire ('Aucun') ;

  if ExistePieceNature('EEX',Depot,DD) then
    begin
    if DateClot<>iDate1900 then
      St:=St+' ; '+RequeteAvecDateCloture(Depot,'gq_entreesorties','EEX',DateClot)
    else if DD<>iDate1900 then
      St:=St+' ; '+RequeteAvecDate2(Depot,'gq_entreesorties','EEX',DD)
    else
      St:=St+' ; '+RequeteSansDateCloture2(Depot,'gq_entreesorties','EEX');
    end
  else St := St + ' ; ' + TraduireMemoire ('Aucun') ;
  TraceLeRecalcul(8,False,St);
end;

procedure TFCalculStock.CalculZoneInv(TT:TQRProgressForm; Depot:string; DateClot:TDateTime);
var DD : TDateTime;
    St : string;
begin
  TraceLeRecalcul(9,True);
  TT.SubText := 'Calcul la zone des écarts d''inventaire'; TT.Value:=TT.Value+1;
  DD:=StrToDate(DtVenteFFO.Text) ;
  if DateClot<>iDate1900 then DD:=DateClot;

  if ExistePieceNature('INV',Depot,DD) then
    begin
    if DateClot<>iDate1900 then
      St:=RequeteAvecDateCloture(Depot,'gq_ecartinv','INV',DateClot)
    else if DD<>iDate1900 then
      St:=RequeteAvecDate(Depot,'gq_ecartinv','INV',DD)
    else
      St:=RequeteSansDateCloture(Depot,'gq_ecartinv','INV');
    end
  else St := TraduireMemoire ('Aucun') ;
  TraceLeRecalcul(9,False,St);
end;

procedure TFCalculStock.CalculZoneTransfert(TT:TQRProgressForm; Depot:string; DateClot:TDateTime);
var DD : TDateTime;
    St : string;
begin
  TraceLeRecalcul(10,True);
  TT.SubText := 'Calcul la zone des transferts'; TT.Value:=TT.Value+1;
  DD:=StrToDate(DtVenteFFO.Text) ;

  if DateClot<>iDate1900 then
    St:=RequeteAvecDateCloture(Depot,'gq_transfert','TEM',DateClot,True)
  else if DD<>iDate1900 then
    St:=RequeteAvecDate(Depot,'gq_transfert','TEM',DD,True)
  else
    St:=RequeteSansDateCloture(Depot,'gq_transfert','TEM',True);

  if DateClot<>iDate1900 then
    St:=St+' ; '+RequeteAvecDateCloture(Depot,'gq_transfert','TRE',DateClot)
  else if DD<>iDate1900 then
    St:=St+' ; '+RequeteAvecDate2(Depot,'gq_transfert','TRE',DD)
  else
    St:=St+' ; '+RequeteSansDateCloture2(Depot,'gq_transfert','TRE');
  TraceLeRecalcul(10,False,St);
end;

procedure TFCalculStock.CalculZoneLivFou(TT:TQRProgressForm; Depot:string; DateClot:TDateTime);
var DD : TDateTime;
    St : string;
begin
  TraceLeRecalcul(11,True);
  TT.SubText := 'Calcul la zone livraison fournisseur'; TT.Value:=TT.Value+1;
  DD:=StrToDate(DtVenteFFO.Text) ;
  if DateClot<>iDate1900 then DD:=DateClot;

  if ExistePieceNature('BLF',Depot,DD) then
    begin
    if DateClot<>iDate1900 then
      St:=RequeteAvecDateCloture(Depot,'gq_livrefou','BLF',DateClot)
    else if DD<>iDate1900 then
      St:=RequeteAvecDate(Depot,'gq_livrefou','BLF',DD)
    else
      St:=RequeteSansDateCloture(Depot,'gq_livrefou','BLF');
    end
  else St := TraduireMemoire ('Aucun') ;
  TraceLeRecalcul(11,False,St);
end;

procedure TFCalculStock.CalculZoneLivCli(TT:TQRProgressForm; Depot:string; DateClot:TDateTime);
var DD : TDateTime;
    St : string;
begin
  TraceLeRecalcul(12,True);
  TT.SubText := 'Calcul la zone livraison client'; TT.Value:=TT.Value+1;
  DD:=StrToDate(DtVenteFFO.Text) ;
  if DateClot<>iDate1900 then DD:=DateClot;

  if ExistePieceNature('BLC',Depot,DD) then
    begin
    if DateClot<>iDate1900 then
      St:=RequeteAvecDateCloture(Depot,'gq_livreclient','BLC',DateClot)
    else if DD<>iDate1900 then
      St:=RequeteAvecDate(Depot,'gq_livreclient','BLC',DD)
    else
      St:=RequeteSansDateCloture(Depot,'gq_livreclient','BLC');
    end
  else St := TraduireMemoire ('Aucun') ;
  TraceLeRecalcul(12,False,St);
end;

procedure TFCalculStock.CalculZonePhysique(TT:TQRProgressForm; Depot:string);
var SQL : string;
    Nb : integer;
begin
  TraceLeRecalcul(13,True);
  TT.SubText := 'Calcul la zone stock physique'; TT.Value:=TT.Value+1;
  SQL := 'Update dispo set gq_physique = '+
         'gq_entreesorties + gq_ecartinv - gq_venteffo - gq_livreclient '+
         '+ gq_livrefou + gq_transfert '+
         'where gq_depot="'+Depot+'" and gq_cloture="-"';
  Nb := ExecuteSQL(SQL);
  TraceLeRecalcul(13,False,IntToStr(Nb));
end;

procedure TFCalculStock.ClotureStockDate(TT:TQRProgressForm; Depot:string);
var SQL : string;
    Nb : integer;
begin
  TraceLeRecalcul(14,True);
  TT.SubText := 'Création du stock clôturé au '+DtClotureDate.Text+' en utilisant le stock courant'; TT.Value:=TT.Value+1;
  SQL := 'INSERT DISPO '+
         'select GQ_ARTICLE,GQ_DEPOT,GQ_STOCKMIN,GQ_STOCKMAX,GQ_DATEINVENTAIRE,'+
         'GQ_PHYSIQUE,GQ_RESERVECLI,GQ_RESERVEFOU,GQ_PREPACLI,GQ_LIVRECLIENT,'+
         'GQ_LIVREFOU,GQ_TRANSFERT,GQ_QTE1,GQ_QTE2,GQ_QTE3,GQ_DPA,GQ_PMAP,'+
         'GQ_DPR,GQ_PMRP,'+
         'GQ_DATECREATION,GQ_DATEMODIF,GQ_DATEINTEGR,GQ_UTILISATEUR,'+
         'GQ_STOCKINITIAL,"X","'+UsDateTime(StrToDate(DtClotureDate.Text))+
         '",GQ_EMPLACEMENT,GQ_CUMULSORTIES,GQ_CUMULENTREES,'+
         'GQ_VENTEFFO,GQ_ENTREESORTIES,GQ_ECARTINV from DISPO '+
         'where gq_depot="'+Depot+'" and gq_cloture="-"';
  Nb := ExecuteSQL(SQL);
  TraceLeRecalcul(14,False,IntToStr(Nb));
end;

procedure TFCalculStock.LanceCalcul;
var ListeDepot,CodeDepot : string;
    TT : TQRProgressForm;
    DateDeCloture : TDateTime;
begin
   TOK.Visible := False;
   if GP_DEPOT.Text='' then
     begin PGIError('La zone dépôt doit-être renseignée.',FCalculStock.Caption); GP_Depot.SetFocus; exit; end;
   ListeDepot := ChargeListeDepot; if ListeDepot='' then exit;

   if (TChkClotureDate.Checked) and (StrToDate(DtClotureDate.Text)=iDate1900) then
     begin PGIError('Vous ne pouvez pas clotûrer le stock à cette date.',FCalculStock.Caption); DtClotureDate.SetFocus; exit; end;

   TraceLeRecalcul(1,True);
   {$IFDEF AGL550B}
     InitMoveProgressForm(FCalculStock,FCalculStock.Caption,'Calcul des stocks.',10,False,True);
   {$ELSE}
     TT:=DebutProgressForm(FCalculStock,FCalculStock.Caption,'Calcul des stocks.',10,False,True);
   {$ENDIF}
   Try
     TT.MaxValue:=NbZonesCochees; TT.Value:=0;
     CodeDepot:=ReadTokenSt(ListeDepot);
     if TChkEgal.Checked then SupEgal1:='>=' else SupEgal1:='>';
     if TChkEgalDate.Checked then SupEgal2:='>=' else SupEgal2:='>';
     if CodeDepot<>'' then
       begin
       if BVerifDepot then VerifDepot(TT,CodeDepot);
       if BVerifStock then VerifStock(TT,CodeDepot);
       end;
     While CodeDepot<>'' do
       begin
       TraceLeRecalcul(2,True,CodeDepot);
       DateDeCloture:=iDate1900;
       TT.Text:='Mise à jour des stocks pour le dépôt : '+RechDom('GCDEPOT',CodeDepot,False);
       if BDateCloture then DateDeCloture := FindDateCloture(TT,CodeDepot);
       if (DateDeCloture<>iDate1900) and BPhysique then InitialiseDispo2(TT,CodeDepot,DateDeCloture);
       if TChkClotureStock.Checked then InitialiseDispo(TT,CodeDepot);
       if BFFO then CalculZoneFFO(TT,CodeDepot,DateDeCloture);
       if BEEX then CalculZoneEEX(TT,CodeDepot,DateDeCloture);
       if BInv then CalculZoneInv(TT,CodeDepot,DateDeCloture);
       if BTransfert then CalculZoneTransfert(TT,CodeDepot,DateDeCloture);
       if BLivFou then CalculZoneLivFou(TT,CodeDepot,DateDeCloture);
       if BLivCli then CalculZoneLivCli(TT,CodeDepot,DateDeCloture);
       {if BPrepaLiv then ;
       if BCommFou then ;
       if BCommCli then ;}
       if BPhysique then CalculZonePhysique(TT,CodeDepot);
       if BClotureDate then ClotureStockDate(TT,CodeDepot);
       CodeDepot:=ReadTokenSt(ListeDepot);
       end;
     TraceLeRecalcul(15,True);
     TOK.Visible := True;
   Finally
     TT.Free;
     end;
end;

procedure TFCalculStock.BloqueEtSelectionZonesAZero(BPhys:boolean);
begin
  TChkPhysique_0.Enabled:=BPhys;
  TChkVenteFFO_0.Enabled:=BPhys;
  TChkEntreesSorties_0.Enabled:=BPhys;
  TChkEcartINV_0.Enabled:=BPhys;
  TChkTransfert_0.Enabled:=BPhys;
  TChkLivreFou_0.Enabled:=BPhys;
  TChkLivreClient_0.Enabled:=BPhys;
end;

procedure TFCalculStock.BloqueEtSelectionZonesStock(BPhys,Force:boolean);
var Bloque : Boolean;
begin
  if Not BPhys then
    begin
    TChkVenteFFO.Checked:=True;
    TChkEntreesSorties.Checked:=True;
    TChkEcartINV.Checked:=True;
    TChkTransfert.Checked:=True;
    TChkLivreFou.Checked:=True;
    TChkLivreClient.Checked:=True;
    {TChkPrepaCli.Checked:=True;
    TChkReserveFou.Checked:=True;
    TchkReserveCli.Checked:=True;}
    TChkEgalDate.Checked:=True;
    end;
  Bloque := BPhys or Force;
  TChkVenteFFO.Enabled:=Bloque; DtVenteFFO.Enabled:=Bloque;
  TChkEntreesSorties.Enabled:=Bloque; DtEntreesSorties.Enabled:=Bloque;
  TChkEcartINV.Enabled:=Bloque; DtEcartINV.Enabled:=Bloque;
  TChkTransfert.Enabled:=Bloque; DtTransfert.Enabled:=Bloque;
  TChkLivreFou.Enabled:=Bloque; DtLivreFou.Enabled:=Bloque;
  TChkLivreClient.Enabled:=Bloque; DtLivreClient.Enabled:=Bloque;
  {TChkPrepaCli.Enabled:=Bloque; DtPrepaCli.Enabled:=Bloque;
  TChkReserveFou.Enabled:=Bloque; DtReserveFou.Enabled:=Bloque;
  TchkReserveCli.Enabled:=Bloque; DtReserveCli.Enabled:=Bloque;}
  TChkEgalDate.Enabled:=Bloque;
end;

procedure TFCalculStock.TChkPhysiqueClick(Sender: TObject);
begin
  inherited;
  BloqueEtSelectionZonesStock(Not TChkPhysique.Checked,V_PGI.SAV);
end;

procedure TFCalculStock.TChkClotureStockClick(Sender: TObject);
begin
  inherited;
  BloqueEtSelectionZonesAZero(Not TChkClotureStock.Checked);
end;

end.
