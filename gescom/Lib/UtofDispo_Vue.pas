unit UtofDispo_Vue;

interface

uses  StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
      HCtrls, HEnt1, HMsgBox, UTOF, M3FP, Stat, HQry,
{$IFDEF EAGLCLIENT}
      emul, MaineAGL,
{$ELSE}
      dbTables, db, DBGrids, mul, Fe_main,
{$ENDIF}
      UTobView, TVProp, UTob, UtilArticle, EntGC ,UtilGC, utilPGI;

Type
     TOF_DISPO_VUE = Class (TOF)
     private
        TobViewer1: TTobViewer;
        stChampsCompl : string ;  // Champs complémentaires de la requête
        procedure TVOnDblClickCell(Sender: TObject ) ;
        Procedure SetLastError (Num : integer; ou : string ) ;
        Function Formate_StWhereVente (NaturesDoc : string) : String;
     public
        StatVente : String ; // Statistique portant sur les ventes d'une période
        ListeNatPiece, DatePieceDeb, DatePieceFin : String ;
        procedure OnArgument(Arguments : String ) ; override ;
        procedure OnLoad    ; override ;
        procedure OnUpdate  ; override ;
     END ;

const
	// libellés des messages
	TexteMessage: array[1..3] of string 	= (
          {1}        'Cet article est défini en taille unique'
          {2}       ,'Opération impossible'
          {3}       ,'Attention, vous n''avez pas sélectionné de nature de document'
          ) ;


implementation


procedure TOF_DISPO_VUE.SetLastError (Num : integer; ou : string );
begin
if ou<>'' then SetFocusControl(ou);
LastError:=Num;
LastErrorMsg:=TexteMessage[LastError];
PGIBox(LastErrorMsg,TexteMessage[2]) ;
end ;

procedure TOF_DISPO_VUE.OnArgument(Arguments : String ) ;
var Nbr,iCol : integer ;
    stIndice : string ;
begin
inherited ;
// Paramétrage des libellés des familles, collection, stat. article et dimensions
ChangeLibre2('TGA_COLLECTION',Ecran);
for iCol:=1 to 3 do
    begin
    stIndice:=IntToStr(iCol) ;
    ChangeLibre2('TGA_FAMILLENIV'+stIndice,Ecran);
    end;
stChampsCompl:='' ;
if (ctxMode in V_PGI.PGIContexte) and (GetPresentation=ART_ORLI) then
    begin
    for iCol:=4 to 8 do
        begin
        stIndice:=IntToStr(iCol) ;
        if ChangeLibre2('TGA2_FAMILLENIV'+stIndice,Ecran) then stChampsCompl:=stChampsCompl+',GA2_FAMILLENIV'+stIndice ;
        end;
    for iCol:=1 to 2 do
        begin
        stIndice:=IntToStr(iCol) ;
        if ChangeLibre2('TGA2_STATART'+stIndice,Ecran) then stChampsCompl:=stChampsCompl+',GA2_STATART'+stIndice ;
        end;
    end ;

TobViewer1:=TTobViewer(getcontrol('TV'));
TobViewer1.OnDblClick:= TVOnDblClickCell ;

if (ctxMode in V_PGI.PGIContexte) then SetControlProperty('GQ_DEPOT','Plus','GDE_SURSITE="X"');

// Paramètrage des libellés des onglets des zones et tables libres du dépôt
if not VH_GC.GCMultiDepots then
   begin
   SetControlCaption('PTABLESLIBRESDEP','Tables Libres Etab.') ;
   SetControlCaption('PZONESLIBRESDEP','Zones Libres Etab.') ;
   end;

// Paramétrage des libellés des tables libres
if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_LIBREART', 10, '') = 0) then SetControlVisible('PTABLESLIBRES', False) ;
if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GDE_LIBREDEP', 10, '') = 0) then SetControlVisible('PTABLESLIBRESDEP', False) ;

// Mise en forme des libellés des dates, booléans libres et montants libres
Nbr := 0;
if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False) else inc(Nbr) ;
if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GA_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False) else inc(Nbr) ;
if (GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'GA_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False) else inc(Nbr) ;
{$IFNDEF CCS3}
if (Nbr = 0) then
{$ENDIF}
   SetControlVisible('PZONESLIBRES',False) ;
Nbr := 0;
if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GDE_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VALDEP', False) else inc(Nbr) ;
if (GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GDE_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATEDEP', False) else inc(Nbr) ;
if (GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'GDE_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOLDEP', False) else inc(Nbr) ;
{$IFNDEF CCS3}
if (Nbr = 0) then
{$ENDIF}
   SetControlVisible('PZONESLIBRESDEP',False) ;
// Initialisation de l'onglet : Compléments
SetControlChecked ('VAL_DPA', False);
SetControlChecked ('VAL_DPR', False);
SetControlChecked ('VAL_PMAP', True);
SetControlChecked ('VAL_PMRP', False);
SetControlChecked ('STAT_VENTE', False);
if (VH_GC.GCSeria = True) or (V_PGI.VersionDemo = True)
   then SetControlText('GL_NATUREPIECEG','FAC;AVC;FFO')
   else SetControlText('GL_NATUREPIECEG','FFO');

StatVente := '-';
ListeNatPiece:='' ;
DatePieceDeb:='' ;
DatePieceFin:='' ;
end;

procedure TOF_DISPO_VUE.OnLoad;
var xx_where : string ;
begin
xx_where:='' ;

// Gestion des checkBox : booléens libres article
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'BOOL', 'GA_BOOLLIBRE', 3, '');

// Gestion des dates libres article
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'DATE', 'GA_DATELIBRE', 3, '_');

// Gestion des montants libres article
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'EDIT', 'GA_VALLIBRE', 3, '_');

// Gestion des checkBox : booléens libres  dépôts
  xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'BOOL', 'GDE_BOOLLIBRE', 3, '');

// Gestion des dates libres dépôts
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'DATE', 'GDE_DATELIBRE', 3, '_');

// Gestion des montants libres dépôts
xx_where:=GCXXWhereChampLibre (TForm(Ecran), xx_where, 'EDIT', 'GDE_VALLIBRE', 3, '_');

SetControlText('XX_WHERE',xx_where) ;

// Contrôle sur l'onglet : Compléments
if (GetControlText('STAT_VENTE')='X') then
   begin
   if (GetControlText('GL_NATUREPIECEG')='') or (GetControlText('GL_NATUREPIECEG')='<<Tous>>') then
      begin
      SetLastError(3, 'GL_NATUREPIECEG'); exit ;
      end;
   end;
end;

procedure TOF_DISPO_VUE.OnUpdate ;
var StSQL, StWhere : String ;
    NaturesDoc : String ;
    StWhereVte : String ;
begin
  Inherited ;
  StatVente := GetControlText('STAT_VENTE');
  if (StatVente='X') then
     begin
     // Statistique portant sur les ventes d'une période
     NaturesDoc   := GetControlText('GL_NATUREPIECEG');
     DatePieceDeb := GetControlText('GL_DATEPIECED');
     DatePieceFin := GetControlText('GL_DATEPIECEF');
     StWhereVte := Formate_StWhereVente (NaturesDoc);
     end;

  // Sauvegarde de la clause WHERE générée automatiquement par l'intermédiaire
  // du MUL.
  StWhere := TFStat(ecran).StSQL;

  // Constitution de la requête SQL
  StSQL := 'SELECT GQ_DEPOT,GQ_CLOTURE,ART.GA_STATUTART,ART.GA_CODEARTICLE,ART.GA_LIBELLE,ART.GA_COLLECTION,'+
           'ART.GA_FOURNPRINC,FOU.T_LIBELLE as NOM_FOURNISSEUR,GQ_EMPLACEMENT,' +
           'ART.GA_FAMILLENIV1,ART.GA_FAMILLENIV2,ART.GA_FAMILLENIV3' ;
  if (ctxMode in V_PGI.PGIContexte) and (GetPresentation=ART_ORLI)
      then StSQL := StSQL + stChampsCompl ;

  StSQL := StSQL + ',(SUM (GQ_LIVREFOU)) QTE_RECU_FOURN,' +
           '(SUM (GQ_TRANSFERT)) QTE_TRANSFERE,' +
           '(SUM (GQ_ENTREESORTIES)) QTE_ENTREE_SORTIE,' +
           '(SUM (GQ_PHYSIQUE)) QTE_STOCK,' +
           '(SUM (GQ_RESERVECLI)) QTE_RESA_CLIENT,' +
           '(SUM (GQ_RESERVEFOU)) QTE_CDE_FOURN,' +
           '(SUM (GQ_PHYSIQUE-GQ_RESERVECLI+GQ_RESERVEFOU)) QTE_STOCK_NET,';

  if (StatVente='X') then
     begin
     StSQL := StSQL + '(Select SUM(GL_QTEFACT) from LIGNE ' + StWhereVte + ') QTE_VENDU,' +
              '(Select SUM(GL_TOTALHT) from LIGNE ' + StWhereVte + ') MTHT_VENDU,' +
              '(Select SUM(GL_TOTALTTC) from LIGNE ' + StWhereVte + ') MTTC_VENDU';
     end else
     begin
     StSQL := StSQL + '(SUM (GQ_LIVRECLIENT)) QTE_LIVRE_CLIENT,' +
                      '(SUM (GQ_VENTEFFO)) QTE_VENTE_FO,' +
                      '(SUM (GQ_LIVRECLIENT+GQ_VENTEFFO)) QTE_VENDU';
     end;

  if (GetControlText('VAL_DPA')<>'-') then
     begin
     if (StatVente='X') then
          StSQL := StSQL + ',(Select SUM(GL_QTEFACT*GL_DPA) from LIGNE ' + StWhereVte + ') DPA_VENDU'
     else StSQL := StSQL + ',(SUM ((GQ_LIVRECLIENT+GQ_VENTEFFO)*GQ_DPA)) DPA_VENDU';

     StSQL := StSQL + ',(SUM (GQ_LIVREFOU*GQ_DPA)) DPA_RECU_FOURN,' +
           '(SUM (GQ_TRANSFERT*GQ_DPA)) DPA_TRANSFERE,' +
           '(SUM (GQ_ENTREESORTIES*GQ_DPA)) DPA_ENTREE_SORTIE,' +
           '(SUM (GQ_PHYSIQUE*GQ_DPA)) DPA_STOCK,' +
           '(SUM (GQ_RESERVECLI*GQ_DPA)) DPA_RESA_CLIENT,' +
           '(SUM (GQ_RESERVEFOU*GQ_DPA)) DPA_CDE_FOURN,' +
           '(SUM ((GQ_PHYSIQUE-GQ_RESERVECLI+GQ_RESERVEFOU)*GQ_DPA)) DPA_STOCK_NET';
     end;

  if (GetControlText('VAL_PMAP')<>'-') then
     begin
     if (StatVente='X') then
          StSQL := StSQL + ',(Select SUM(GL_QTEFACT*GL_PMAP) from LIGNE ' + StWhereVte + ') PMAP_VENDU'
     else StSQL := StSQL + ',(SUM ((GQ_LIVRECLIENT+GQ_VENTEFFO)*GQ_PMAP)) PMAP_VENDU';

     StSQL := StSQL + ',(SUM (GQ_LIVREFOU*GQ_PMAP)) PMAP_RECU_FOURN,' +
           '(SUM (GQ_TRANSFERT*GQ_PMAP)) PMAP_TRANSFERE,' +
           '(SUM (GQ_ENTREESORTIES*GQ_PMAP)) PMAP_ENTREE_SORTIE,' +
           '(SUM (GQ_PHYSIQUE*GQ_PMAP)) PMAP_STOCK,' +
           '(SUM (GQ_RESERVECLI*GQ_PMAP)) PMAP_RESA_CLIENT,' +
           '(SUM (GQ_RESERVEFOU*GQ_PMAP)) PMAP_CDE_FOURN,' +
           '(SUM ((GQ_PHYSIQUE-GQ_RESERVECLI+GQ_RESERVEFOU)*GQ_PMAP)) PMAP_STOCK_NET';
     end;

  if (GetControlText('VAL_DPR')<>'-') then
     begin
     if (StatVente='X') then
          StSQL := StSQL + ',(Select SUM(GL_QTEFACT*GL_DPA) from LIGNE ' + StWhereVte + ') DPR_VENDU'
     else StSQL := StSQL + ',(SUM ((GQ_LIVRECLIENT+GQ_VENTEFFO)*GQ_DPR)) DPR_VENDU';

     StSQL := StSQL + ',(SUM (GQ_LIVREFOU*GQ_DPR)) DPR_RECU_FOURN,' +
           '(SUM (GQ_TRANSFERT*GQ_DPR)) DPR_TRANSFERE,' +
           '(SUM (GQ_ENTREESORTIES*GQ_DPR)) DPR_ENTREE_SORTIE,' +
           '(SUM (GQ_PHYSIQUE*GQ_DPR)) DPR_STOCK,' +
           '(SUM (GQ_RESERVECLI*GQ_DPR)) DPR_RESA_CLIENT,' +
           '(SUM (GQ_RESERVEFOU*GQ_DPR)) DPR_CDE_FOURN,' +
           '(SUM ((GQ_PHYSIQUE-GQ_RESERVECLI+GQ_RESERVEFOU)*GQ_DPR)) DPR_STOCK_NET';
     end;

  if (GetControlText('VAL_PMRP')<>'-') then
     begin
     if (StatVente='X') then
          StSQL := StSQL + ',(Select SUM(GL_QTEFACT*GL_DPA) from LIGNE ' + StWhereVte + ') PMRP_VENDU'
     else StSQL := StSQL + ',(SUM ((GQ_LIVRECLIENT+GQ_VENTEFFO)*GQ_PMRP)) PMRP_VENDU';

     StSQL := StSQL + ',(SUM (GQ_LIVREFOU*GQ_PMRP)) PMRP_RECU_FOURN,' +
           '(SUM (GQ_TRANSFERT*GQ_PMRP)) PMRP_TRANSFERE,' +
           '(SUM (GQ_ENTREESORTIES*GQ_PMRP)) PMRP_ENTREE_SORTIE,' +
           '(SUM (GQ_PHYSIQUE*GQ_PMRP)) PMRP_STOCK,' +
           '(SUM (GQ_RESERVECLI*GQ_PMRP)) PMRP_RESA_CLIENT,' +
           '(SUM (GQ_RESERVEFOU*GQ_PMRP)) PMRP_CDE_FOURN,' +
           '(SUM ((GQ_PHYSIQUE-GQ_RESERVECLI+GQ_RESERVEFOU)*GQ_PMRP)) PMRP_STOCK_NET';
     end;

  StSQL := StSQL + ' FROM DISPO' +
            ' LEFT JOIN DEPOTS ON GDE_DEPOT=GQ_DEPOT' +
            ' LEFT JOIN ARTICLE ART ON GQ_ARTICLE=GA_ARTICLE' ;
  if (ctxMode in V_PGI.PGIContexte) and (GetPresentation=ART_ORLI)
      then StSQL := StSQL + ' LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=GQ_ARTICLE' ;
  StSQL := StSQL + ' LEFT JOIN TIERS FOU ON ART.GA_FOURNPRINC=FOU.T_AUXILIAIRE';
  StSQL := StSQL + ' WHERE GQ_CLOTURE="-" AND ART.GA_TYPEARTICLE="MAR" AND ART.GA_STATUTART<>"GEN"';
  StSQL := StSQL + ' GROUP BY GQ_DEPOT,GQ_CLOTURE,ART.GA_STATUTART,ART.GA_CODEARTICLE,ART.GA_LIBELLE,' +
           'ART.GA_COLLECTION,ART.GA_FOURNPRINC,FOU.T_LIBELLE,ART.GA_FAMILLENIV1,' +
           'ART.GA_FAMILLENIV2,ART.GA_FAMILLENIV3,GQ_EMPLACEMENT';
  if (ctxMode in V_PGI.PGIContexte) and (GetPresentation=ART_ORLI)
      then StSQL := StSQL + stChampsCompl ;

  // Ajout du complément de la clause WHERE, pour prendre en compte les critères
  // de sélection renseignés par l'utilisateur (collection, fournisseur, ...)
  stWhere := RecupWhereCritere(TPageControl(GetControl('Pages'))) ;
  if Uppercase(Copy(StWhere,1,6))='WHERE ' then
     begin
     StWhere := Copy(StWhere,7,length(StWhere)-6) ;
     StSQL := InsertSQLWhere(StSQL, StWhere) ;
     end;
  TFStat(ecran).StSQL := StSQL;
end ;

Function TOF_DISPO_VUE.Formate_StWhereVente(NaturesDoc : String) : String;
var StWhereVte : String ;
    CodeNature : string;
begin
StWhereVte := 'Where ';
ListeNatPiece := '';
if NaturesDoc<>'<<Tous>>' then
   begin
   While NaturesDoc<>'' do
     begin
     CodeNature:=ReadTokenSt(NaturesDoc);
     if ListeNatPiece<>'' then ListeNatPiece := ListeNatPiece+',';
     ListeNatPiece := ListeNatPiece+'"'+CodeNature+'"';
     end;
   end;

if ListeNatPiece<>'' then
   begin
   StWhereVte := StWhereVte + 'GL_NATUREPIECEG IN ('+ListeNatPiece+') and ';
   end;
StWhereVte := StWhereVte + 'GL_DATEPIECE>="'+UsDateTime(StrToDate(DatePieceDeb))+'" and ' +
              'GL_DATEPIECE<="'+UsDateTime(StrToDate(DatePieceFin))+'" and ' +
              'GL_DEPOT=GQ_DEPOT and GL_CODEARTICLE=ART.GA_CODEARTICLE and ' +
              'GL_TYPELIGNE="ART"';
result := StWhereVte;
end;

procedure TOF_DISPO_VUE.TVOnDblClickCell(Sender: TObject );
var bCloture : boolean;
    st1, st2, depot, cloture, statutart, args : string ;
begin
with TTobViewer(sender) do
    begin
    st1:=AsString[ColIndex('GA_CODEARTICLE'), CurrentRow] ;
    st2 := CodeArticleUnique2(st1, '') ;
    depot:=AsString[ColIndex('GQ_DEPOT'), CurrentRow] ;
    bCloture:=AsBoolean[ColIndex('GQ_CLOTURE'), CurrentRow] ;
    cloture:=CheckToString(bCloture) ;
    statutart:=AsString[ColIndex('GA_STATUTART'), CurrentRow] ;
    if (statutart<>'GEN') and (statutart<>'DIM')
    then begin SetLastError(1, '') ; exit end
    else if (StatVente='X') then args:='ACTION=CONSULTATION;STOCKVTE=X;DEPOT='+depot+
               ';CLOTURE='+cloture+';NATPIECE='+ListeNatPiece+';DATEPIECEDEB='+DatePieceDeb+
               ';DATEPIECEFIN='+DatePieceFin
    else args:='ACTION=CONSULTATION;STOCK=X;DEPOT='+depot+';CLOTURE='+cloture ;
    V_PGI.DispatchTT (7,taConsult,st2,args,'') ;     // inversion OT 06/09/02

{         AGLLanceFiche('GC','GCARTICLE', '', st2, 'ACTION=CONSULTATION;STOCKVTE=X;DEPOT='+depot+
               ';CLOTURE='+cloture+';NATPIECE='+ListeNatPiece+';DATEPIECEDEB='+DatePieceDeb+
               ';DATEPIECEFIN='+DatePieceFin)
    else AGLLanceFiche('GC','GCARTICLE', '', st2, 'ACTION=CONSULTATION;STOCK=X;DEPOT='+depot+';CLOTURE='+cloture) ;
}

    //if (ColName[CurrentCol] = 'RPE_AUXILIAIRE') or (ColName[CurrentCol] = 'T_LIBELLE') then
    //  AGLLanceFiche('GC','GCTIERS','',AsString[ColIndex('RPE_AUXILIAIRE'), CurrentRow],'MONOFICHE;T_NATUREAUXI='+AsString[ColIndex('RPE_TYPETIERS'), CurrentRow])
    //else if (ColName[CurrentCol] = 'RPE_INTERVENANT') or (ColName[CurrentCol] = 'GCL_LIBELLE') then
    //  AGLLanceFiche('GC','GCCOMMERCIAL','',AsString[ColIndex('RPE_INTERVENANT'), CurrentRow],'ACTION=MODIFICATION')
    //else if (ColName[CurrentCol] = 'RPE_NUMACTION') then
    // AGLLanceFiche('RT','RTACTIONS','',AsString[ColIndex('RPE_AUXILIAIRE'), CurrentRow]+';'+IntToStr(AsInteger[ColIndex('RPE_NUMACTION'), CurrentRow]),'ACTION=CONSULTATION')
    //else
    // AGLLanceFiche('RT','RTPERSPECTIVES','',IntToStr(AsInteger[ColIndex('RPE_PERSPECTIVE'), CurrentRow]) ,'');
    end;
end;

Initialization
registerclasses([TOF_DISPO_VUE]);

end.
