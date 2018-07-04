unit UTofMBOAffMinMax;

interface
uses  M3FP,StdCtrls,Controls,Classes,Dialogs,saisutil,HRichOLE,Spin,
      HStatus,
      HCtrls,HEnt1,HMsgBox, Hdimension, UTOB, UTOF, AGLInit,EntGC,Buttons,
{$IFDEF EAGLCLIENT}
      Maineagl,
{$ELSE}
      dbTables,DBGrids, db,FE_Main,
{$ENDIF}
      forms,sysutils,ComCtrls,
      HDB, vierge, Math ;

Type
     TOF_MBOAFFMINMAX = Class (TOF)
     public
        NbFicheTot,NbArtTot : integer;
        TobListeDepot,TobDepotSel : TOB ;
        TobArticle,TobArtUniq : TOB ;
        TobDispo : TOB ;
        GLISTE,GAFFICHE : THGrid ;
        procedure OnClose ; override ;
        Procedure OnLoad  ; override ;
        procedure OnArgument (Arguments : String ) ; override ;
        procedure OnMajStkBtq;
        procedure ClickFlecheDroite ;
        procedure ClickFlecheGauche ;
        procedure ClickFlecheTous ;
        procedure ClickFlecheAucun ;
        procedure ClickFlecheHaut ;
        procedure ClickFlecheBas ;
        procedure RefreshGrid(posListe,posAffiche : integer) ;
        procedure RefreshBouton ;
        Procedure ChargeSelectBtq ;
        procedure TraitementMinMax(TobArtMul:TOB; ListeDepotSel:string; StkMin, StkMax:double);
        function  FindStInTAB(St:String;Tab:Array of string):boolean ;
        procedure MajTableDISPO() ;
     END ;

Const
    BTN_DROIT   = 'DROIT' ;
    BTN_GAUCHE  = 'GAUCHE' ;
    BTN_HAUT    = 'HAUT' ;
    BTN_BAS     = 'BAS' ;
    BTN_TOUS    = 'TOUS' ;
    BTN_AUCUN   = 'AUCUN' ;
    GRD_LISTE   = 'GLISTE' ;
    GRD_AFFICHE = 'GAFFICHE' ;


implementation

Const NbArtParRequete : integer = 200;

Procedure TOF_MBOAFFMINMAX.OnArgument(Arguments : String ) ;
begin
inherited ;
GAFFICHE:=THGRID(GetControl('GAFFICHE'));
GAFFICHE.ColWidths[0]:=161 ; GAFFICHE.ColAligns[0]:=taLeftJustify ;
GLISTE:=THGRID(GetControl('GLISTE'));
GLISTE.ColWidths[0]:=161 ; GLISTE.ColAligns[0]:=taLeftJustify ;
end ;

Procedure TOF_MBOAFFMINMAX.OnLoad  ;
var Q : TQuery;
    TypeAff : THValComboBox ;

BEGIN
inherited ;
TobListeDepot:=TOB.CREATE('Liste établissements',NIL,-1) ;
TobDepotSel:=TOB.CREATE('Etablissements affichés',NIL,-1) ;
Q:=OpenSQL('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE = "GAS"',False) ;
if Not Q.EOF then
  begin
  TypeAff:=THValComboBox(GetControl('TYPEAFFECTATION')) ;
  TypeAff.Value:=Q.Fields[0].AsString ; // Initialisation par défaut
  end;
Ferme(Q) ;
ChargeSelectBtq
END;

procedure TOF_MBOAFFMINMAX.OnClose ;
begin
inherited ;
TobDepotSel.free;
TobDepotSel:=nil;
TobListeDepot.free;
TobListeDepot:=nil;
end;

procedure TOF_MBOAFFMINMAX.RefreshGrid(posListe,posAffiche : integer) ;
begin
TobDepotSel.PutGridDetail(GAFFICHE,False,False,'ET_LIBELLE',True) ;
TobListeDepot.PutGridDetail(GLISTE,False,False,'ET_LIBELLE',True) ;
GAFFICHE.Row:=Min(posAffiche,GAFFICHE.RowCount-1) ;
GLISTE.Row:=Min(posListe,GLISTE.RowCount-1) ;
RefreshBouton ;
end ;

// Boutons enable / disable
procedure TOF_MBOAFFMINMAX.RefreshBouton ;
begin
SetControlEnabled('BFLECHEDROITE',TobListeDepot.Detail.Count>0) ;
SetControlEnabled('BFLECHEGAUCHE',TobDepotSel.Detail.Count>0) ;
SetControlEnabled('BFLECHEHAUT',GAFFICHE.Row>1) ;
SetControlEnabled('BFLECHEBAS',GAFFICHE.Row<GAFFICHE.RowCount-1) ;
SetControlEnabled('BFLECHETOUS',TobListeDepot.Detail.Count>0) ;
SetControlEnabled('BFLECHEAUCUN',TobDepotSel.Detail.Count>0) ;
end ;

procedure TOF_MBOAFFMINMAX.ClickFlecheDroite ;
var indiceFille : integer ;
begin
// Y a t il quelque chose de sélectionné ?
if GLISTE.Row<0 then exit ;
// Changement du parent de l'élément de la liste des établissements
if TobDepotSel.Detail.Count>0 then indiceFille:=GAFFICHE.Row else indiceFille:=0 ;
TobListeDepot.detail[GLISTE.Row-1].ChangeParent(TobDepotSel,indiceFille) ;
RefreshGrid(GLISTE.Row,GAFFICHE.Row+1) ;
end ;

procedure TOF_MBOAFFMINMAX.ClickFlecheGauche ;
var indiceFille : integer ;
begin
// Y a t il quelque chose de sélectionné ?
if GAFFICHE.Row<0 then exit ;
// Changement du parent de l'élément des établissements affichés
if TobListeDepot.Detail.Count>0 then indiceFille:=GLISTE.Row else indiceFille:=0 ;
TobDepotSel.detail[GAFFICHE.Row-1].ChangeParent(TobListeDepot,indiceFille) ;
RefreshGrid(GLISTE.Row+1,GAFFICHE.Row) ;
end ;

procedure TOF_MBOAFFMINMAX.ClickFlecheTous ;
var indiceFille,iGrd, posliste : integer ;
begin
if GLISTE.RowCount<2 then exit ;
// Changement du parent de l'élément de la liste des établissements
if GAFFICHE.RowCount>2 then indiceFille:=GAFFICHE.Row else indiceFille:=0 ;
posliste := TobListeDepot.detail.count-1;
for iGrd:=0 to posliste do TobListeDepot.detail[0].ChangeParent(TobDepotSel,indiceFille+iGrd) ;
RefreshGrid(1,indiceFille+posliste) ;
end ;

procedure TOF_MBOAFFMINMAX.ClickFlecheAucun ;
var indiceFille,iGrd, posaffiche : integer ;
begin
if GAFFICHE.RowCount<2 then exit ;
// Changement du parent de l'élément de la liste des établissements
posaffiche := TobDepotSel.detail.count-1;
if GLISTE.RowCount>2 then indiceFille:=GLISTE.Row else indiceFille:=0 ;
for iGrd:=0 to posaffiche do TobDepotSel.detail[0].ChangeParent(TobListeDepot,indiceFille+iGrd) ;
RefreshGrid(indiceFille+posaffiche,1) ;
end ;

procedure TOF_MBOAFFMINMAX.ClickFlecheHaut ;
begin
if GAFFICHE.Row<1 then exit ;
// Changement de l'indice dans la Tob parent
TobDepotSel.detail[GAFFICHE.Row-1].ChangeParent(TobDepotSel,GAFFICHE.Row-2) ;
RefreshGrid(GLISTE.Row,GAFFICHE.Row-1) ;
end ;

procedure TOF_MBOAFFMINMAX.ClickFlecheBas ;
begin
if GAFFICHE.Row>GAFFICHE.RowCount-2 then exit ;
// Changement de l'indice dans la Tob parent
TobDepotSel.detail[GAFFICHE.Row-1].ChangeParent(TobDepotSel,GAFFICHE.Row) ;
RefreshGrid(GLISTE.Row,GAFFICHE.Row+1) ;
end ;

Procedure TOF_MBOAFFMINMAX.ChargeSelectBtq;
var QQ : TQuery ;
begin
inherited ;
if not (ctxMode in V_PGI.PGIContexte) then exit ;
// Chargement de la liste des établissements et du paramétrage existant.
TobListeDepot.ClearDetail ;
TobDepotSel.ClearDetail ;
// Chargement paramétrage sauvegardé
QQ:=OpenSQL('select ET_ETABLISSEMENT,ET_LIBELLE from ETABLISS order by ET_ETABLISSEMENT',True) ;
if not QQ.eof then TobListeDepot.LoadDetailDB('','','',QQ,false);
Ferme(QQ) ;

// Affichage des tobs
GLISTE:=THGrid(GetControl('GLISTE')) ;
GAFFICHE:=THGrid(GetControl('GAFFICHE')) ;
TobDepotSel.PutGridDetail(GAFFICHE,False,False,'ET_LIBELLE',True) ;
TobListeDepot.PutGridDetail(GLISTE,False,False,'ET_LIBELLE',True) ;
RefreshBouton ;
end;

Procedure AGLOnMajStkBtq (Parms: array of variant; nb: integer) ;
var F : TForm;
     TOTOF : TOF ;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge)
    then TOTOF := TFVierge(F).LaTof
    else exit;
  if (TOTOF is TOF_MBOAFFMINMAX) then TOF_MBOAFFMINMAX(TOTOF).OnMajStkBtq else exit;
end;

procedure TOF_MBOAFFMINMAX.OnMajStkBtq;
var i_ind : integer;
    ListeDepotSel : string;
    TobArtMul : TOB;
    RecapMaj :THRichEditOle;
    lancetrait: TBitBtn;
    StkMin, StkMax : double;

begin
StkMin := TSpinEdit(GetControl('QTESTOCKMIN')).Value;
StkMax := TSpinEdit(GetControl('QTESTOCKMAX')).Value;
if (StkMin <> 0) or (StkMax <> 0) then
  if StkMin >= StkMax then
    begin
    MessageAlerte('La quantité maxi doit être supérieure à la quantité mini');
    exit;
    end;

// Tob contenant les articles sélectionnés depuis le MUL Articles
TobArtMul := LaTOB ;
LaTOB := Nil ;

// Formattage d'une chaine de caractères contenant la liste des dépôts sélectionnés.
ListeDepotSel := '';
for i_ind:=0 to TobDepotSel.detail.count-1 do
  begin
  if ListeDepotSel <> '' then ListeDepotSel := ListeDepotSel+',';
  ListeDepotSel := ListeDepotSel+'"'+TobDepotSel.Detail[i_ind].GetValue('ET_ETABLISSEMENT')+'"';
  end;
if ListeDepotSel = '' then
  begin
  MessageAlerte('Pas de boutique à traiter ');
  exit;
  end;

lancetrait:=TBitBtn(GetControl('BGENERE'));
lancetrait.enabled := False;
RecapMaj:=THRichEditOle(GetControl('OLE1'));
RecapMaj.Lines.Text := 'Traitement en cours ...';
NbFicheTot := 0;
NbArtTot := 0;

try
  TobArticle:=TOB.CREATE('Article Dim',NIL,-1) ;
  TobArtUniq:=TOB.CREATE('Article Uniq',NIL,-1) ;
  TobDispo:=TOB.CREATE('Dispo Art',NIL,-1) ;

  // Lancement du traitement de mise à jour des stocks mini et maxi
  TraitementMinMax(TobArtMul, ListeDepotSel, StkMin, StkMax);
  finally
  TobArticle.free;
  TobArtUniq.free;
  TobDispo.free;
  TobArtMul.free;
  end;

if NbArtTot=0
  then RecapMaj.Lines.Add('Pas d''article correspondant aux masques paramétrés')
  else RecapMaj.Lines.Add('Mise à jour de '+inttostr(NbFicheTot)+' fiches stock articles dimensionnés');
RecapMaj.Lines.Add('Traitement terminé ');
lancetrait.enabled := True;
end;

procedure TOF_MBOAFFMINMAX.TraitementMinMax(TobArtMul:TOB; ListeDepotSel:string; StkMin, StkMax:double);
var i,CountStArt,NbRequete,y : integer;
    NbArt : integer;
    StArticle : string;
    SQL : string ;
    TabWhereArticle,TabWhereArtUniq,TabWhereDispo : array of String;
    TOBA : TOB;
    QArticle,QArtUniq,QDispo : TQuery ;
begin
// Constitution d'une ou plusieurs listes d'articles sélectionnés (maximum 200 articles),
// qui alimenteront la clause WHERE des différentes requêtes :
//   - Requête regroupant tous les articles dimensionnés avec la Qté mini et maxi, en
//     fonction du masque de chacun.
//   - Requête regroupant tous les articles unique avec la Qté mini et maxi saisie.
//   - Requête regroupant le contenu de la table DISPO pour tous les articles sélectionnés.
CountStArt:=0; NbRequete:=0;
SetLength(TabWhereArticle, 1); TabWhereArticle[0]:='';
SetLength(TabWhereArtUniq, 1); TabWhereArtUniq[0]:='';
SetLength(TabWhereDispo, 1);   TabWhereDispo[0]:='';
INITMOVE(TobArtMul.Detail.Count,'');
For i:=0 to TobArtMul.Detail.count-1 do
  begin
  MoveCur(False);
  TOBA:=TobArtMul.Detail[i];
  StArticle:=TOBA.GetValue('GA_CODEARTICLE');
  if CountStArt>=NbArtParRequete then
    begin
    NbRequete:=NbRequete+1;
    SetLength(TabWhereArticle, NbRequete+1); TabWhereArticle[NbRequete]:='';
    SetLength(TabWhereArtUniq, NbRequete+1); TabWhereArtUniq[NbRequete]:='';
    SetLength(TabWhereDispo, NbRequete+1);   TabWhereDispo[NbRequete]:='';
    CountStArt:=0;
    end;
  if Not FindStInTAB(StArticle,TabWhereDispo) then
    begin
    CountStArt:=CountStArt+1;
    if TabWhereDispo[NbRequete]<>'' then TabWhereDispo[NbRequete]:=TabWhereDispo[NbRequete]+',';
    TabWhereDispo[NbRequete] := TabWhereDispo[NbRequete]+'"'+StArticle+'"';

    if TOBA.GetValue('GA_STATUTART')='UNI' then
      begin
      if TabWhereArtUniq[NbRequete]<>'' then TabWhereArtUniq[NbRequete]:=TabWhereArtUniq[NbRequete]+',';
      TabWhereArtUniq[NbRequete] := TabWhereArtUniq[NbRequete]+'"'+StArticle+'"';
      end else
      begin
      if TabWhereArticle[NbRequete]<>'' then TabWhereArticle[NbRequete]:=TabWhereArticle[NbRequete]+',';
      TabWhereArticle[NbRequete] := TabWhereArticle[NbRequete]+'"'+StArticle+'"';
      end;
    end;
  end;
FINIMOVE;

if TabWhereDispo[0]<>'' then
  begin
  for y:=Low(TabWhereDispo) to High(TabWhereDispo) do
    begin
    if TabWhereDispo[y]<>'' then
      begin
      // TOB du DISPO actuel sur les différents dépôts
      SQL:='Select GQ_ARTICLE,GQ_DEPOT,GQ_CLOTURE,GQ_STOCKMIN,GQ_STOCKMAX '+
               'From ARTICLE Left join DISPO on GA_ARTICLE=GQ_ARTICLE '+
               'Where GA_CODEARTICLE IN ('+TabWhereDispo[y]+') and GA_STATUTART<>"GEN"'+
               ' and GQ_DEPOT IN ('+ListeDepotSel+') and GQ_CLOTURE="-"';
      QDispo:=OpenSQL(SQL,True);
      if Not QDispo.EOF then TobDispo.LoadDetailDB('DISPO','','',QDispo,True);
      Ferme(QDispo);
      end;
    if TabWhereArticle[y]<>'' then
      begin
      // TOB reprenant pour chaque article dimensionné, les dernières qté mini et maxi
      // enregistrées en fonction du masque de dimension.
      SQL:='select GA_ARTICLE,GA_CODEARTICLE, GAM_QTEDISPOMINI, GAM_QTEDISPOMAXI from ARTICLE '+
       'left join ARTICLESERIE on GA_DIMMASQUE=GAM_DIMMASQUE '+
       'where GA_STATUTART<>"GEN" and GA_CODEARTICLE IN ('+TabWhereArticle[y]+') and '+
       'GAM_CODESERIE = "'+GetControltext('TYPEAFFECTATION')+'"'+
       ' and (GA_CODEDIM1 is null or GA_CODEDIM1=GAM_CODEDIM1)'+
       ' and (GA_CODEDIM2 is null or GA_CODEDIM2=GAM_CODEDIM2)'+
       ' and (GA_CODEDIM3 is null or GA_CODEDIM3=GAM_CODEDIM3)'+
       ' and (GA_CODEDIM4 is null or GA_CODEDIM4=GAM_CODEDIM4)'+
       ' and (GA_CODEDIM5 is null or GA_CODEDIM5=GAM_CODEDIM5)' ;
      QArticle:=OpenSQL(SQL,True);
      if Not QArticle.EOF then TobArticle.LoadDetailDB('','','',QArticle,True);
      Ferme(QArticle);
      end;

    if (TabWhereArtUniq[y]<>'') and ((StkMin <> 0) or (StkMax <> 0)) then
      begin
      // TOB reprenant pour chaque article unique, la qté mini et maxi saisie
      // Cette Tob est ensuite déversée dans TobArticle.
      SQL:='select GA_ARTICLE,GA_CODEARTICLE from ARTICLE '+
           'where GA_CODEARTICLE IN ('+TabWhereArtUniq[y]+')';
      QArtUniq:=OpenSQL(SQL,True);
      if Not QArtUniq.EOF then TobArtUniq.LoadDetailDB('','','',QArtUniq,True);
      Ferme(QArtUniq);

      while (TobArtUniq.Detail.Count>0) do
        begin
        TOBA:=TobArtUniq.Detail[0];
        TOBA.AddChampSup ('GAM_QTEDISPOMINI', False);
        TOBA.AddChampSup ('GAM_QTEDISPOMAXI', False);
        TOBA.PutValue ('GAM_QTEDISPOMINI', StkMin);
        TOBA.PutValue ('GAM_QTEDISPOMAXI', StkMax);
        TOBA.ChangeParent(TobArticle,-1);
        end;
      end;

      NbArt:=TobArticle.Detail.count;
      NbArtTot:=NbArtTot+NbArt;
      if NbArt>0 then
        begin
        // Mise à jour de la TOB du DISPO et enregistrement de celle-ci dans la base
        MajTableDISPO();
        end;

      TobArticle.ClearDetail;
      TobDispo.ClearDetail;
    end;
  end;
end;

function TOF_MBOAFFMINMAX.FindStInTAB(St:String;Tab:Array of string):boolean;
var i : integer;
begin
result:=false;
for i:=Low(Tab) to High(Tab) do
  if Pos(St,Tab[i])<>0 then begin result:=true; break; end;
end;

// Mise à jour de la TOB du DISPO et enregistrement de celle-ci dans la base
procedure TOF_MBOAFFMINMAX.MajTableDISPO();
var i,j : integer;
    NbFiche : integer;
    Depot : string;
    QteMin,QteMax : double;
    TOBA, TOBD : TOB;
begin
NbFiche:=0;
For i:=0 to TobDepotSel.Detail.count-1 do
  begin
  Depot := TobDepotSel.Detail[i].GetValue('ET_ETABLISSEMENT');
  INITMOVE(TobArticle.Detail.Count,'');
  For j:=0 to TobArticle.Detail.count-1 do
    begin
    MoveCur(False);
    TOBA:=TobArticle.Detail[j];
    QteMin:=TOBA.GetValue('GAM_QTEDISPOMINI');
    QteMax:=TOBA.GetValue('GAM_QTEDISPOMAXI');
    TOBD:=TobDispo.FindFirst(['GQ_ARTICLE','GQ_DEPOT'],[TOBA.GetValue('GA_ARTICLE'),Depot],false);
    if TOBD=Nil then
      begin
      if (QteMin<>0) or (QteMax<>0) then
        begin
        TOBD:=TOB.Create('DISPO',TobDispo,-1);
        TOBD.PutValue('GQ_ARTICLE', TOBA.GetValue('GA_ARTICLE'));
        TOBD.PutValue('GQ_DEPOT', Depot);
        TOBD.PutValue('GQ_CLOTURE', '-');
        TOBD.PutValue('GQ_STOCKMIN', QteMin);
        TOBD.PutValue('GQ_STOCKMAX', QteMax);
        inc (NbFicheTot);
        inc (NbFiche);
        end;
      end else
      begin
      if (QteMin<>TOBD.GetValue('GQ_STOCKMIN')) or (QteMax<>TOBD.GetValue('GQ_STOCKMAX')) then
        begin
        TOBD.PutValue('GQ_STOCKMIN', QteMin);
        TOBD.PutValue('GQ_STOCKMAX', QteMax);
        inc (NbFicheTot);
        inc (NbFiche);
        end;
      end;
    end;
  FINIMOVE;
  end;
// Mise à jour de la table DISPO
if NbFiche>0 then TobDispo.InsertOrUpdateDB(True);
end;

procedure AGLOnClickBouton4 (Parms: array of variant; nb: integer) ;
var F : TForm ;
    TOTOF : TOF ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then TOTOF:=TFVierge(F).LaTOF else exit ;
if (TOTOF is TOF_MBOAFFMINMAX) then
    BEGIN
    if Parms[1]=BTN_DROIT then TOF_MBOAFFMINMAX(TOTOF).ClickFlecheDroite
    else if Parms[1]=BTN_GAUCHE then TOF_MBOAFFMINMAX(TOTOF).ClickFlecheGauche
    else if Parms[1]=BTN_HAUT then TOF_MBOAFFMINMAX(TOTOF).ClickFlecheHaut
    else if Parms[1]=BTN_BAS then TOF_MBOAFFMINMAX(TOTOF).ClickFlecheBas
    else if Parms[1]=BTN_TOUS then TOF_MBOAFFMINMAX(TOTOF).ClickFlecheTous
    else if Parms[1]=BTN_AUCUN then TOF_MBOAFFMINMAX(TOTOF).ClickFlecheAucun
    else if (Parms[1]=GRD_LISTE) or (Parms[1]=GRD_AFFICHE) then TOF_MBOAFFMINMAX(TOTOF).RefreshBouton ;
    END ;
end ;

Initialization
registerclasses([TOF_MBOAFFMINMAX]) ;
RegisterAglProc('OnClickBouton4', True , 1, AGLOnClickBouton4) ;
RegisterAglProc('OnMajStkBtq', True , 1, AGLOnMajStkBtq) ;
end.

