unit EuroPGI;

interface


Uses windows,hent1,sysutils,UTOB,classes,Forms,ExtCtrls,SaisUtil,
{$IFDEF EAGLCLIENT}
{$ELSE}
     db,
     {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
     MajTable,
{$ENDIF}
     HCtrls,Ent1, HQry, UtilPGI;

Function BasculeEuroPGI ( StDev : String ; PanelPGI : TPanel ) : Boolean ;

implementation

Var OldDev : String ;

Procedure InverseChamps ( Q : TDataSet ; C1,C2 : String ) ;
Var X,Y : Double ;
BEGIN
X:=Q.FindField(C1).AsFloat ; Y:=Q.FindField(C2).AsFloat ;
Q.FindField(C1).AsFloat:=Y ; Q.FindField(C2).AsFloat:=X ;
END ;

Procedure SwapLesCons ( Q : TDataSet ; NomChamp : String ; VerifExchamp : boolean = False ) ;
BEGIN
if VerifExChamp then
   BEGIN
   // Because FactAff qui plante partout ...
   if Q.FindField(NomChamp)=Nil then Exit ;
   if Q.FindField(NomChamp+'CON')=Nil then Exit ;
   END ;
InverseChamps(Q,NomChamp,NomChamp+'CON') ;
END ;

Procedure SwapSaisie ( Q : TDataSet ; NomChamp : String ) ;
Var St : String ;
BEGIN
St:=Q.FindField(NomChamp).AsString ;
if St='-' then St:='X' else St:='-' ;
Q.FindField(NomChamp).AsString:=St ;
END ;

Procedure CalcChampEuro ( Q : TDataSet ; NomC : String ) ;
Var TauxE : double;
BEGIN
if (V_PGI.TauxEuro=0) then TauxE:=1 else TauxE:=V_PGI.TauxEuro;
if Q.FindField(NomC+'CON').AsFloat=0 then
   BEGIN
   Q.FindField(NomC+'CON').AsFloat:=Arrondi(Q.FindField(NomC).AsFloat/TauxE,2) ;
   END ;
END ;

Procedure CalculeEuro ( Q : TDataSet ; NomT : String ) ;
BEGIN
if NomT='ARTICLE' then
   BEGIN
   CalcChampEuro(Q,'GA_DPA')  ; CalcChampEuro(Q,'GA_DPR')   ;
   CalcChampEuro(Q,'GA_PMAP') ; CalcChampEuro(Q,'GA_PMRP')  ;
   CalcChampEuro(Q,'GA_PVHT') ; CalcChampEuro(Q,'GA_PVTTC') ;
   CalcChampEuro(Q,'GA_PAHT') ; CalcChampEuro(Q,'GA_PRHT')  ;
   CalcChampEuro(Q,'GA_MARGEMINI') ;
   END else
if NomT='CATALOGU' then
   BEGIN
   CalcChampEuro(Q,'GCA_PRIXBASE') ;
   CalcChampEuro(Q,'GCA_DPA') ;
   END else
if NomT='DISPO' then
   BEGIN
   CalcChampEuro(Q,'GQ_DPA')  ; CalcChampEuro(Q,'GQ_DPR')  ;
   CalcChampEuro(Q,'GQ_PMAP') ; CalcChampEuro(Q,'GQ_PMRP') ;
   END else
if NomT='LIGNENOMEN' then
   BEGIN
   CalcChampEuro(Q,'GLN_DPA')  ; CalcChampEuro(Q,'GLN_DPR')  ;
   CalcChampEuro(Q,'GLN_PMAP') ; CalcChampEuro(Q,'GLN_PMRP') ;
   END else ;
END ;

Procedure ConvertEuroGC ( Q : TDataSet ; NomT : String );
Var i    : integer ;
DEV          : RDEVISE ;
Prix     : DOUBLE ;
BEGIN
if NomT='' then Exit ;
CalculeEuro(Q,NomT) ;
if NomT='PIECE' then
   BEGIN
   SwapLesCons(Q,'GP_TOTALESC')     ; SwapLesCons(Q,'GP_TOTESCTTC') ;
   SwapLesCons(Q,'GP_TOTALREMISE')  ; SwapLesCons(Q,'GP_TOTREMISETTC') ;
   SwapLesCons(Q,'GP_TOTALHT')      ; SwapLesCons(Q,'GP_TOTALTTC') ;
   SwapLesCons(Q,'GP_TOTALBASEREM') ; SwapLesCons(Q,'GP_TOTALBASEESC') ;
   SwapLesCons(Q,'GP_ACOMPTE') ;
   if Q.FindField('GP_DEVISE').AsString=OldDev then
      BEGIN
      SwapSaisie(Q,'GP_SAISIECONTRE') ;
      Q.FindField('GP_DEVISE').AsString:='EUR' ;
      END ;
   END else
if NomT='PIEDBASE' then
   BEGIN
   InverseChamps(Q,'GPB_BASETAXE','GPB_BASECON') ;
   InverseChamps(Q,'GPB_VALEURTAXE','GPB_VALEURCON') ;
   if Q.FindField('GPB_DEVISE').AsString=OldDev then
      BEGIN
      SwapSaisie(Q,'GPB_SAISIECONTRE') ;
      Q.FindField('GPB_DEVISE').AsString:='EUR' ;
      END ;
   END else
if NomT='PIEDECHE' then
   BEGIN
   InverseChamps(Q,'GPE_MONTANTECHE','GPE_MONTANTCON') ;
   if Q.FindField('GPE_DEVISE').AsString=OldDev then
      BEGIN
      SwapSaisie(Q,'GPE_SAISIECONTRE') ;
      Q.FindField('GPE_DEVISE').AsString:='EUR' ;
      END ;
   END else
if NomT='PIEDPORT' then
   BEGIN
   SwapLesCons(Q,'GPT_BASEHT')  ; SwapLesCons(Q,'GPT_BASETTC') ;
   SwapLesCons(Q,'GPT_TOTALHT') ; SwapLesCons(Q,'GPT_TOTALTTC') ;
   for i:=1 to 5 do InverseChamps(Q,'GPT_TOTALTAXE'+IntToStr(i),'GPT_TOTALTAXECON'+IntToStr(i)) ;
   if Q.FindField('GPT_DEVISE').AsString=OldDev then
      BEGIN
      Q.FindField('GPT_DEVISE').AsString:='EUR' ;
      END ;
   Q.FindField('GPT_MONTANTMINI').AsFloat:=FrancToEuro(Q.FindField('GPT_MONTANTMINI').AsFloat) ;
   Q.FindField('GPT_MINIMUM').AsFloat:=FrancToEuro(Q.FindField('GPT_MINIMUM').AsFloat) ;
   END else
if NomT='ACOMPTES' then
   BEGIN
   SwapLesCons(Q,'GAC_MONTANT') ;
   END else
if NomT='VENTANA' then
   BEGIN
   SwapLesCons(Q,'YVA_MONTANT') ;
   END else
if NomT='LIGNE' then
   BEGIN
   SwapLesCons(Q,'GL_DPA')         ; SwapLesCons(Q,'GL_DPR')  ;
   SwapLesCons(Q,'GL_PMAP')        ; SwapLesCons(Q,'GL_PMAPACTU')  ;
   SwapLesCons(Q,'GL_PMRP')        ; SwapLesCons(Q,'GL_PMRPACTU') ;
   SwapLesCons(Q,'GL_TOTALHT')     ; SwapLesCons(Q,'GL_TOTALTTC') ;
   SwapLesCons(Q,'GL_MONTANTHT')   ; SwapLesCons(Q,'GL_MONTANTTTC') ;
   SwapLesCons(Q,'GL_TOTREMPIED')  ; SwapLesCons(Q,'GL_TOTREMLIGNE') ;
   SwapLesCons(Q,'GL_PUHT')        ; SwapLesCons(Q,'GL_PUTTC') ;
   SwapLesCons(Q,'GL_PUHTNET')     ; SwapLesCons(Q,'GL_PUTTCNET') ;
   SwapLesCons(Q,'GL_TOTESCLIGNE') ;
   for i:=1 to 5 do InverseChamps(Q,'GL_TOTALTAXE'+IntToStr(i),'GL_TOTALTAXECON'+IntToStr(i)) ;
   if Q.FindField('GL_DEVISE').AsString=OldDev then
      BEGIN
      Q.FindField('GL_DEVISE').AsString:='EUR' ;
      SwapSaisie(Q,'GL_SAISIECONTRE') ;
      END ;
   END else
if NomT='TARIF' then
   BEGIN
   if Q.FindField('GF_DEVISE').AsString=OldDev then
      BEGIN
      //InverseChamps(Q,'GF_PRIXUNITAIRE','GF_PRIXCON') ;
      //Q.FindField('GF_DEVISE').AsString:='EUR' ;
      Q.FindField('GF_PRIXCON').AsFloat:=Q.FindField('GF_PRIXUNITAIRE').AsFloat ;
      END
      else
      BEGIN
      // Conversion du tarif dans la devise fongible
         // Conversion du tarif en Euro
      DEV.Code:=Q.FindField('GF_DEVISE').AsString ;
      DEV.DateTaux := NowH ;
      GetInfosDevise(DEV) ;
      DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, Date) ;
      Prix:=Q.FindField('GF_PRIXUNITAIRE').AsFloat ;
      Prix:=DeviseToEuro (Prix, DEV.Taux, DEV.Quotite);
          // Conversion du tarif en DeviseFongible
      DEV.Code:=OldDev ;
      DEV.DateTaux := NowH ;
      GetInfosDevise(DEV) ;
      DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, Date) ;
      Prix:=EuroToDevise (Prix,DEV.Taux,DEV.Quotite,DEV.Decimale) ;
      Q.FindField('GF_PRIXCON').AsFloat:=Prix ;
      END ;
   END else
if NomT='ARTICLE' then
   BEGIN
   SwapLesCons(Q,'GA_DPA')  ; SwapLesCons(Q,'GA_DPR') ;
   SwapLesCons(Q,'GA_PVHT') ; SwapLesCons(Q,'GA_PVTTC') ;
   SwapLesCons(Q,'GA_PAHT') ; SwapLesCons(Q,'GA_PRHT') ;
   SwapLesCons(Q,'GA_PMAP') ; SwapLesCons(Q,'GA_PMRP') ;
   END else
if NomT='CATALOGU' then
   BEGIN
   SwapLesCons(Q,'GCA_DPA') ; SwapLesCons(Q,'GCA_PRIXBASE') ;
   Q.FindField('GCA_PRIXVENTE').AsFloat:=FrancToEuro(Q.FindField('GCA_PRIXVENTE').AsFloat) ;
   END else
if NomT='LIGNENOMEN' then
   BEGIN
   SwapLesCons(Q,'GLN_DPA')  ; SwapLesCons(Q,'GLN_DPR') ;
   SwapLesCons(Q,'GLN_PMAP') ; SwapLesCons(Q,'GLN_PMRP') ;
   END else
if NomT='PORT' then
   BEGIN
   Q.FindField('GPO_PVHT').AsFloat:=FrancToEuro(Q.FindField('GPO_PVHT').AsFloat) ;
   Q.FindField('GPO_PVTTC').AsFloat:=FrancToEuro(Q.FindField('GPO_PVTTC').AsFloat) ;
   Q.FindField('GPO_MINIMUM').AsFloat:=FrancToEuro(Q.FindField('GPO_MINIMUM').AsFloat) ;
   Q.FindField('GPO_MINIMUMTTC').AsFloat:=FrancToEuro(Q.FindField('GPO_MINIMUMTTC').AsFloat) ;
   END else
if NomT='PRIXREVIENT' then
   BEGIN
   if Q.FindField('PRV_TYPECALCUL').AsString='MON' then Q.FindField('PRV_VALEUR').AsFloat:=FrancToEuro(Q.FindField('PRV_VALEUR').AsFloat) ;
   END else
if NomT='TIERS' then
   BEGIN
   Q.FindField('T_TOTDERNPIECE').AsFloat:=FrancToEuro(Q.FindField('T_TOTDERNPIECE').AsFloat) ;
   Q.FindField('T_CREDITACCORDE').AsFloat:=FrancToEuro(Q.FindField('T_CREDITACCORDE').AsFloat) ;
   Q.FindField('T_CREDITPLAFOND').AsFloat:=FrancToEuro(Q.FindField('T_CREDITPLAFOND').AsFloat) ;
   Q.FindField('T_CREDITDEMANDE').AsFloat:=FrancToEuro(Q.FindField('T_CREDITDEMANDE').AsFloat) ;
   END else
if NomT='DISPO' then
   BEGIN
   SwapLesCons(Q,'GQ_DPA') ; SwapLesCons(Q,'GQ_PMAP') ;
   SwapLesCons(Q,'GQ_DPR') ; SwapLesCons(Q,'GQ_PMRP') ;
   END else
if NomT='ACTIVITE' then
   BEGIN
   Q.FindField('ACT_TOTPR').AsFloat:=FrancToEuro(Q.FindField('ACT_TOTPR').AsFloat) ;
   Q.FindField('ACT_TOTPRCHARGE').AsFloat:=FrancToEuro(Q.FindField('ACT_TOTPRCHARGE').AsFloat) ;
   Q.FindField('ACT_TOTPRCHINDI').AsFloat:=FrancToEuro(Q.FindField('ACT_TOTPRCHINDI').AsFloat) ;
   SwapLesCons(Q,'ACT_TOTVENTE')  ;
   if Q.FindField('ACT_DEVISE').AsString=OldDev then
      BEGIN
      Q.FindField('ACT_DEVISE').AsString:='EUR' ;
      END ;
   // ajout MD 26/04/01
   Q.FindField('ACT_PUPR').AsFloat:=FrancToEuro(Q.FindField('ACT_PUPR').AsFloat) ;
   Q.FindField('ACT_PUPRCHARGE').AsFloat:=FrancToEuro(Q.FindField('ACT_PUPRCHARGE').AsFloat) ;
   Q.FindField('ACT_PUPRCHINDIRECT').AsFloat:=FrancToEuro(Q.FindField('ACT_PUPRCHINDIRECT').AsFloat) ;
   SwapLesCons(Q,'ACT_PUVENTE')  ;
   // Ajout PA le 23/08/2001
   Q.FindField('ACT_PUVENTEDEV').AsFloat := Q.FindField('ACT_PUVENTE').AsFloat;
   Q.FindField('ACT_TOTVENTEDEV').AsFloat:= Q.FindField('ACT_TOTVENTE').AsFloat;
   END else
if NomT='AFFAIRE' then
   BEGIN
   SwapLesCons(Q,'AFF_TOTALHT')    ; SwapLesCons(Q,'AFF_TOTALTTC') ;
   SwapLesCons(Q,'AFF_REPORT')     ; SwapLesCons(Q,'AFF_TOTALTAXE') ;
   SwapLesCons(Q,'AFF_TOTALHTGLO') ; SwapLesCons(Q,'AFF_MONTANTECHE') ;
   if Q.FindField('AFF_DEVISE').AsString=OldDev then
      BEGIN
      SwapSaisie(Q,'AFF_SAISIECONTRE') ;
      Q.FindField('AFF_DEVISE').AsString:='EUR' ;
      END ;
   END else
if NomT='AFFTIERS' then
   BEGIN
   Q.FindField('AFT_PROPOINTERV').AsFloat:=FrancToEuro(Q.FindField('AFT_PROPOINTERV').AsFloat) ;
   END else
if NomT='FACTAFF' then
   BEGIN
   SwapLesCons(Q,'AFA_MONTANTECHE',True) ; SwapLesCons(Q,'AFA_BONIMALI',True) ;
   SwapLesCons(Q,'AFA_AFACTURER',True)   ; SwapLesCons(Q,'AFA_BM1FO',True) ;
   SwapLesCons(Q,'AFA_BM1FR',True) ;
   SwapLesCons(Q,'AFA_BM2PR',True)       ; SwapLesCons(Q,'AFA_BM2FO',True);
   SwapLesCons(Q,'AFA_BM2FR',True)       ; SwapLesCons(Q,'AFA_AFACTOT',True);
   SwapLesCons(Q,'AFA_AFACTFR',True)     ; SwapLesCons(Q,'AFA_AFACTFO',True);
   SwapLesCons(Q,'AFA_ACPTEPR',True)     ; SwapLesCons(Q,'AFA_ACPTEFO',True);
   SwapLesCons(Q,'AFA_ACPTEFR',True)     ;
   if Q.FindField('AFA_DEVISE').AsString=OldDev then
      BEGIN
      Q.FindField('AFA_DEVISE').AsString:='EUR' ;
      END ;
   END else
if NomT='FONCTION' then
   BEGIN
   Q.FindField('AFO_TAUXREVIENTUN').AsFloat:=FrancToEuro(Q.FindField('AFO_TAUXREVIENTUN').AsFloat) ;
   Q.FindField('AFO_PVHT').AsFloat:=FrancToEuro(Q.FindField('AFO_PVHT').AsFloat) ;
   Q.FindField('AFO_PVTTC').AsFloat:=FrancToEuro(Q.FindField('AFO_PVTTC').AsFloat) ;
   END else
if NomT='HISTOACTIVITE' then
   BEGIN
   Q.FindField('AHI_TOTPRPRES').AsFloat:=FrancToEuro(Q.FindField('AHI_TOTPRPRES').AsFloat) ;
   Q.FindField('AHI_TOTPRFRAI').AsFloat:=FrancToEuro(Q.FindField('AHI_TOTPRFRAI').AsFloat) ;
   Q.FindField('AHI_TOTPRFOUR').AsFloat:=FrancToEuro(Q.FindField('AHI_TOTPRFOUR').AsFloat) ;
   Q.FindField('AHI_TOTPVPRES').AsFloat:=FrancToEuro(Q.FindField('AHI_TOTPVPRES').AsFloat) ;
   Q.FindField('AHI_TOTPVFRAI').AsFloat:=FrancToEuro(Q.FindField('AHI_TOTPVFRAI').AsFloat) ;
   Q.FindField('AHI_TOTPVFOUR').AsFloat:=FrancToEuro(Q.FindField('AHI_TOTPVFOUR').AsFloat) ;
   Q.FindField('AHI_TOTFACTPRES').AsFloat:=FrancToEuro(Q.FindField('AHI_TOTFACTPRES').AsFloat) ;
   Q.FindField('AHI_TOTFACTFRAI').AsFloat:=FrancToEuro(Q.FindField('AHI_TOTFACTFRAI').AsFloat) ;
   Q.FindField('AHI_TOTFACTFOUR').AsFloat:=FrancToEuro(Q.FindField('AHI_TOTFACTFOUR').AsFloat) ;
   Q.FindField('AHI_TOTFACTGLOB').AsFloat:=FrancToEuro(Q.FindField('AHI_TOTFACTGLOB').AsFloat) ;
   END else
if NomT='PROFILGENER' then
   BEGIN
   Q.FindField('APG_SEUILMAXI').AsFloat:=FrancToEuro(Q.FindField('APG_SEUILMAXI').AsFloat) ;
   Q.FindField('APG_SEUILMINI').AsFloat:=FrancToEuro(Q.FindField('APG_SEUILMINI').AsFloat) ;
   END else
if NomT='RESSOURCE' then
   BEGIN
   Q.FindField('ARS_TAUXUNIT').AsFloat:=FrancToEuro(Q.FindField('ARS_TAUXUNIT').AsFloat) ;
   Q.FindField('ARS_TAUXREVIENTUN').AsFloat:=FrancToEuro(Q.FindField('ARS_TAUXREVIENTUN').AsFloat) ;
   Q.FindField('ARS_PVHT').AsFloat:=FrancToEuro(Q.FindField('ARS_PVHT').AsFloat) ;
   Q.FindField('ARS_PVTTC').AsFloat:=FrancToEuro(Q.FindField('ARS_PVTTC').AsFloat) ;
   Q.FindField('ARS_PVHTCALCUL').AsFloat:=FrancToEuro(Q.FindField('ARS_PVHTCALCUL').AsFloat) ;
   // Ajout historique des prix PA le 23/08/2001
   for i:=2 to 4 do
      BEGIN
      Q.FindField('ARS_TAUXREVIENTUN'+IntToStr(i)).AsFloat:=FrancToEuro(Q.FindField('ARS_TAUXREVIENTUN'+IntToStr(i)).AsFloat) ;
      Q.FindField('ARS_PVHT'+IntToStr(i)).AsFloat:=FrancToEuro(Q.FindField('ARS_PVHT'+IntToStr(i)).AsFloat) ;
      Q.FindField('ARS_PVTTC'+IntToStr(i)).AsFloat:=FrancToEuro(Q.FindField('ARS_PVTTC'+IntToStr(i)).AsFloat) ;
      Q.FindField('ARS_PVHTCALCUL'+IntToStr(i)).AsFloat:=FrancToEuro(Q.FindField('ARS_PVHTCALCUL'+IntToStr(i)).AsFloat) ;
      END;
   END else
   ;
END ;

Procedure AfficheTitre ( NomTable : String ; PanelPGI : TPanel ) ;
Var Domaine : String ;
    NumT    : integer ;
BEGIN
if PanelPGI=Nil then Exit ;
NumT:=TableToNum(NomTable) ; if NumT<=0 then Exit ;
Domaine:=V_PGI.DeTables[NumT].Domaine ;
if Domaine='G' then
   BEGIN
   PanelPGI.Caption:='Bascule Gestion Commerciale : '+NomTable ;
   END else
if Domaine='A' then
   BEGIN
   PanelPGI.Caption:='Bascule Gestion Affaire : '+NomTable ;
   END else
   ;
Application.ProcessMessages ;
END ;

Function CalculeLigneEuro : String ;
Var TauxE : double;
BEGIN
if (V_PGI.TauxEuro=0) then TauxE:=1 else TauxE:=V_PGI.TauxEuro;
Result:='GL_DPACON=ROUND(GL_DPA/'+FloatToStr(TauxE)+',2), GL_DPRCON=ROUND(GL_DPR/'+FloatToStr(TauxE)+',2), '
       +'GL_PMAPCON=ROUND(GL_PMAP/'+FloatToStr(TauxE)+',2), GL_PMRPCON=ROUND(GL_PMRP/'+FloatToStr(TauxE)+',2), '
       +'GL_PMAPACTUCON=ROUND(GL_PMAPACTU/'+FloatToStr(TauxE)+',2), GL_PMRPACTUCON=ROUND(GL_PMRPACTU/'+FloatToStr(TauxE)+',2) ' ;
//Result:='GL_DPACON=ROUND(GL_DPA/6.55957,2), GL_DPRCON=ROUND(GL_DPR/6.55957,2), '
//       +'GL_PMAPCON=ROUND(GL_PMAP/6.55957,2), GL_PMRPCON=ROUND(GL_PMRP/6.55957,2), '
//       +'GL_PMAPACTUCON=ROUND(GL_PMAPACTU/6.55957,2), GL_PMRPACTUCON=ROUND(GL_PMRPACTU/6.55957,2) ' ;
END ;

Function EchangeLigne ( NomTable : String )  : String ;
BEGIN
if NomTable='LIGNE' then
   Result:='GL_DPA=GL_DPACON,GL_DPACON=GL_DPA, GL_DPR=GL_DPRCON,GL_DPRCON=GL_DPR, '
       +'GL_PMAP=GL_PMAPCON,GL_PMAPCON=GL_PMAP, GL_PMAPACTU=GL_PMAPACTUCON,GL_PMAPACTUCON=GL_PMAPACTU, '
       +'GL_PMRP=GL_PMRPCON,GL_PMRPCON=GL_PMRP, GL_PMRPACTU=GL_PMRPACTUCON,GL_PMRPACTUCON=GL_PMRPACTU, '
       +'GL_TOTALHT=GL_TOTALHTCON,GL_TOTALHTCON=GL_TOTALHT, GL_TOTALTTC=GL_TOTALTTCCON,GL_TOTALTTCCON=GL_TOTALTTC, '
       +'GL_MONTANTHT=GL_MONTANTHTCON,GL_MONTANTHTCON=GL_MONTANTHT, GL_MONTANTTTC=GL_MONTANTTTCCON,GL_MONTANTTTCCON=GL_MONTANTTTC, '
       +'GL_TOTREMPIED=GL_TOTREMPIEDCON,GL_TOTREMPIEDCON=GL_TOTREMPIED, GL_TOTREMLIGNE=GL_TOTREMLIGNECON,GL_TOTREMLIGNECON=GL_TOTREMLIGNE, '
       +'GL_PUHT=GL_PUHTCON,GL_PUHTCON=GL_PUHT, GL_PUTTC=GL_PUTTCCON,GL_PUTTCCON=GL_PUTTC, '
       +'GL_PUHTNET=GL_PUHTNETCON,GL_PUHTNETCON=GL_PUHTNET, GL_PUTTCNET=GL_PUTTCNETCON,GL_PUTTCNETCON=GL_PUTTCNET, '
       +'GL_TOTESCLIGNE=GL_TOTESCLIGNECON,GL_TOTESCLIGNECON=GL_TOTESCLIGNE, '
       +'GL_TOTALTAXE1=GL_TOTALTAXECON1, GL_TOTALTAXECON1=GL_TOTALTAXE1, '
       +'GL_TOTALTAXE2=GL_TOTALTAXECON2, GL_TOTALTAXECON2=GL_TOTALTAXE2, '
       +'GL_TOTALTAXE3=GL_TOTALTAXECON3, GL_TOTALTAXECON3=GL_TOTALTAXE3, '
       +'GL_TOTALTAXE4=GL_TOTALTAXECON4, GL_TOTALTAXECON4=GL_TOTALTAXE4, '
       +'GL_TOTALTAXE5=GL_TOTALTAXECON5, GL_TOTALTAXECON5=GL_TOTALTAXE5 '
else if NomTable='PIECE' then
   Result:='GP_TOTALESC=GP_TOTALESCCON,GP_TOTALESCCON=GP_TOTALESC, '
       +'GP_TOTESCTTC=GP_TOTESCTTCCON,GP_TOTESCTTCCON=GP_TOTESCTTC, GP_TOTALREMISE=GP_TOTALREMISECON,GP_TOTALREMISECON=GP_TOTALREMISE, '
       +'GP_TOTREMISETTC=GP_TOTREMISETTCCON,GP_TOTREMISETTCCON=GP_TOTREMISETTC, GP_TOTALHT=GP_TOTALHTCON,GP_TOTALHTCON=GP_TOTALHT, '
       +'GP_TOTALTTC=GP_TOTALTTCCON,GP_TOTALTTCCON=GP_TOTALTTC, GP_TOTALBASEREM=GP_TOTALBASEREMCON,GP_TOTALBASEREMCON=GP_TOTALBASEREM, '
       +'GP_TOTALBASEESC=GP_TOTALBASEESCCON,GP_TOTALBASEESCCON=GP_TOTALBASEESC, GP_ACOMPTE=GP_ACOMPTECON,GP_ACOMPTECON=GP_ACOMPTE'
else if NomTable='PIEDBASE' then
   Result:='GPB_BASETAXE=GPB_BASECON,GPB_BASECON=GPB_BASETAXE, '
       +'GPB_VALEURTAXE=GPB_VALEURCON,GPB_VALEURCON=GPB_VALEURTAXE '
else if NomTable='PIEDECHE' then
   Result:='GPE_MONTANTECHE=GPE_MONTANTCON,GPE_MONTANTCON=GPE_MONTANTECHE ';
END ;

Procedure UpdateLigneEuro ( NomTable : String ; PanelPGI : TPanel ; StNat,StSouche : String ; Num1,Num2 : Integer) ;
BEGIN
if NomTable='LIGNE' then
   BEGIN
   ExecuteSQL('UPDATE LIGNE SET '+CalculeLigneEuro+' WHERE GL_NATUREPIECEG="'+StNat+'" AND GL_SOUCHE="'+StSouche+'" '
             +'AND GL_NUMERO>='+IntToStr(Num1)+' AND GL_NUMERO<'+IntToStr(Num2)) ;
   ExecuteSQL('UPDATE LIGNE SET '+EchangeLigne(NomTable)+' WHERE GL_NATUREPIECEG="'+StNat+'" AND GL_SOUCHE="'+StSouche+'" '
             +'AND GL_NUMERO>='+IntToStr(Num1)+' AND GL_NUMERO<'+IntToStr(Num2)) ;
   ExecuteSQL('UPDATE LIGNE SET GL_DEVISE="EUR", GL_SAISIECONTRE="-" WHERE GL_NATUREPIECEG="'+StNat+'" AND GL_SOUCHE="'+StSouche+'" '
             +'AND GL_NUMERO>='+IntToStr(Num1)+' AND GL_NUMERO<'+IntToStr(Num2)+' '
             +'AND GL_DEVISE="'+OldDev+'" AND GL_SAISIECONTRE="X"') ;
   ExecuteSQL('UPDATE LIGNE SET GL_DEVISE="EUR", GL_SAISIECONTRE="X" WHERE GL_NATUREPIECEG="'+StNat+'" AND GL_SOUCHE="'+StSouche+'" '
             +'AND GL_NUMERO>='+IntToStr(Num1)+' AND GL_NUMERO<'+IntToStr(Num2)+' '
             +'AND GL_DEVISE="'+OldDev+'" AND GL_SAISIECONTRE="-"') ;
   END
else if NomTable='PIECE' then
   BEGIN
   ExecuteSQL('UPDATE PIECE SET '+EchangeLigne(NomTable)+' WHERE GP_NATUREPIECEG="'+StNat+'" AND GP_SOUCHE="'+StSouche+'" '
             +'AND GP_NUMERO>='+IntToStr(Num1)+' AND GP_NUMERO<'+IntToStr(Num2)) ;
   ExecuteSQL('UPDATE PIECE SET GP_DEVISE="EUR", GP_SAISIECONTRE="-" WHERE GP_NATUREPIECEG="'+StNat+'" AND GP_SOUCHE="'+StSouche+'" '
             +'AND GP_NUMERO>='+IntToStr(Num1)+' AND GP_NUMERO<'+IntToStr(Num2)+' '
             +'AND GP_DEVISE="'+OldDev+'" AND GP_SAISIECONTRE="X"') ;
   ExecuteSQL('UPDATE PIECE SET GP_DEVISE="EUR", GP_SAISIECONTRE="X" WHERE GP_NATUREPIECEG="'+StNat+'" AND GP_SOUCHE="'+StSouche+'" '
             +'AND GP_NUMERO>='+IntToStr(Num1)+' AND GP_NUMERO<'+IntToStr(Num2)+' '
             +'AND GP_DEVISE="'+OldDev+'" AND GP_SAISIECONTRE="-"') ;
   END
else if NomTable='PIEDBASE' then
   BEGIN
   ExecuteSQL('UPDATE PIEDBASE SET '+EchangeLigne(NomTable)+' WHERE GPB_NATUREPIECEG="'+StNat+'" AND GPB_SOUCHE="'+StSouche+'" '
             +'AND GPB_NUMERO>='+IntToStr(Num1)+' AND GPB_NUMERO<'+IntToStr(Num2)) ;
   ExecuteSQL('UPDATE PIEDBASE SET GPB_DEVISE="EUR", GPB_SAISIECONTRE="-" WHERE GPB_NATUREPIECEG="'+StNat+'" AND GPB_SOUCHE="'+StSouche+'" '
             +'AND GPB_NUMERO>='+IntToStr(Num1)+' AND GPB_NUMERO<'+IntToStr(Num2)+' '
             +'AND GPB_DEVISE="'+OldDev+'" AND GPB_SAISIECONTRE="X"') ;
   ExecuteSQL('UPDATE PIEDBASE SET GPB_DEVISE="EUR", GPB_SAISIECONTRE="X" WHERE GPB_NATUREPIECEG="'+StNat+'" AND GPB_SOUCHE="'+StSouche+'" '
             +'AND GPB_NUMERO>='+IntToStr(Num1)+' AND GPB_NUMERO<'+IntToStr(Num2)+' '
             +'AND GPB_DEVISE="'+OldDev+'" AND GPB_SAISIECONTRE="-"') ;
   END
else if NomTable='PIEDECHE' then
   BEGIN
   ExecuteSQL('UPDATE PIEDECHE SET '+EchangeLigne(NomTable)+' WHERE GPE_NATUREPIECEG="'+StNat+'" AND GPE_SOUCHE="'+StSouche+'" '
             +'AND GPE_NUMERO>='+IntToStr(Num1)+' AND GPE_NUMERO<'+IntToStr(Num2)) ;
   ExecuteSQL('UPDATE PIEDECHE SET GPE_DEVISE="EUR", GPE_SAISIECONTRE="-" WHERE GPE_NATUREPIECEG="'+StNat+'" AND GPE_SOUCHE="'+StSouche+'" '
             +'AND GPE_NUMERO>='+IntToStr(Num1)+' AND GPE_NUMERO<'+IntToStr(Num2)+' '
             +'AND GPE_DEVISE="'+OldDev+'" AND GPE_SAISIECONTRE="X"') ;
   ExecuteSQL('UPDATE PIEDECHE SET GPE_DEVISE="EUR", GPE_SAISIECONTRE="X" WHERE GPE_NATUREPIECEG="'+StNat+'" AND GPE_SOUCHE="'+StSouche+'" '
             +'AND GPE_NUMERO>='+IntToStr(Num1)+' AND GPE_NUMERO<'+IntToStr(Num2)+' '
             +'AND GPE_DEVISE="'+OldDev+'" AND GPE_SAISIECONTRE="-"') ;
   END;
END ;

Procedure ConvertLigneEuro ( StNat,StSouche : String ; Num1,Num2 : integer ) ;
BEGIN
ExecuteSQL('UPDATE LIGNE SET '+CalculeLigneEuro+' WHERE GL_NATUREPIECEG="'+StNat+'" AND GL_SOUCHE="'+StSouche+'" '
         +'AND GL_NUMERO>='+IntToStr(Num1)+' AND GL_NUMERO<'+IntToStr(Num2)) ;
END ;

Procedure ConvertTablePGI ( NomTable : String ; PanelPGI : TPanel ; StNat : String = '' ; StSouche : String = '' ;
                            Num1 : Integer = 0 ; Num2 : Integer = 0) ;
Var TT : THTable ;
    ii  : integer ;
    Pref : String ;
BEGIN
AfficheTitre(NomTable,PanelPGI) ; Pref:=TableToPrefixe(NomTable) ;
if NomTable='LIGNE' then ConvertLigneEuro(StNat,StSouche,Num1,Num2) ;
BeginTrans ;
TT:=OpenTable(NomTable,Pref+'_CLE1',False) ;
ii:=0 ;
While Not TT.EOF do
   BEGIN
   TT.Edit ; ConvertEuroGC(TT,NomTable) ; TT.Post ;
   inc(ii) ;
   if ii Mod 500=0 then BEGIN CommitTrans ; BeginTrans ; END ;
   TT.Next ;
   END ;
Ferme(TT) ;
CommitTrans ;
END ;

Procedure ConvertTablePGIPiece ( NomTable : String ; PanelPGI : TPanel) ;
Var NInc,NCou,NMin,NMax,NN1,NN2,NbStep  : integer ;
    ListeNat,ListeSouche : TStrings ;
    Q,QNu : TQuery ;
    i,k : integer ;
BEGIN
// Lecture des natures
ListeNat:=TStringList.Create ; ListeSouche:=TStringList.Create ;
Q:=OpenSQL('Select DISTINCT GP_NATUREPIECEG, GP_SOUCHE FROM PIECE',True) ;
While Not Q.EOF do
      BEGIN
      ListeNat.Add(Q.Fields[0].AsString) ;
      ListeSouche.Add(Q.Fields[1].AsString) ;
      Q.Next ;
      END ;
Ferme(Q) ;
// Balayage des pièces avec découpe
for i:=0 to ListeNat.Count-1 do
    BEGIN
    QNu:=OpenSQL('SELECT MIN(GP_NUMERO), MAX(GP_NUMERO), Count(*) FROM PIECE WHERE GP_NATUREPIECEG="'+ListeNat[i]+'" AND GP_SOUCHE="'+ListeSouche[i]+'"',True) ;
    NMin:=QNu.Fields[0].AsInteger ; NMax:=QNu.Fields[1].AsInteger ; NCou:=QNu.Fields[2].AsInteger ;
    Ferme(QNu) ;
    NbStep:=(NCou div 1000)+1 ;  NInc:=(NMax-NMin) div NbStep ;
    for k:=1 to NbStep do
        BEGIN
        NN1:=NMin+(k-1)*NInc ;
        if k<NbStep then NN2:=NN1+NInc else NN2:=NMax+1 ;
        if V_PGI.Driver=DBMsSQL then UpdateLigneEuro('PIECE',PanelPGI,ListeNat[i],ListeSouche[i],NN1,NN2)
                                else ConvertTablePGI('PIECE',PanelPGI,ListeNat[i],ListeSouche[i],NN1,NN2) ;
        if V_PGI.Driver=DBMsSQL then UpdateLigneEuro('LIGNE',PanelPGI,ListeNat[i],ListeSouche[i],NN1,NN2)
                                else ConvertTablePGI('LIGNE',PanelPGI,ListeNat[i],ListeSouche[i],NN1,NN2) ;
        if V_PGI.Driver=DBMsSQL then UpdateLigneEuro('PIEDECHE',PanelPGI,ListeNat[i],ListeSouche[i],NN1,NN2)
                                else ConvertTablePGI('PIEDECHE',PanelPGI,ListeNat[i],ListeSouche[i],NN1,NN2) ;
        if V_PGI.Driver=DBMsSQL then UpdateLigneEuro('PIEDBASE',PanelPGI,ListeNat[i],ListeSouche[i],NN1,NN2)
                                else ConvertTablePGI('PIEDBASE',PanelPGI,ListeNat[i],ListeSouche[i],NN1,NN2) ;
        END ;
    END ;
ListeNat.Clear ; ListeNat.Free ;
ListeSouche.Clear ; ListeSouche.Free ;
END ;

Function BasculeEuroGescom ( PanelPGI : TPanel ) : boolean ;
BEGIN
Result:=False ;
{Gescom}
ConvertTablePGIPiece('PIECE',PanelPGI) ;
//ConvertTablePGI('PIEDBASE',PanelPGI) ;
//ConvertTablePGI('PIEDECHE',PanelPGI) ;
//ConvertTablePGIPiece('LIGNE',PanelPGI) ;
ConvertTablePGI('PIEDPORT',PanelPGI) ;
ConvertTablePGI('ACOMPTES',PanelPGI) ;
ConvertTablePGI('VENTANA',PanelPGI) ;
ConvertTablePGI('TARIF',PanelPGI) ;
ConvertTablePGI('ARTICLE',PanelPGI) ;
ConvertTablePGI('CATALOGU',PanelPGI) ;
ConvertTablePGI('LIGNENOMEN',PanelPGI) ;
ConvertTablePGI('PORT',PanelPGI) ;
ConvertTablePGI('PRIXREVIENT',PanelPGI) ;
ConvertTablePGI('TIERS',PanelPGI) ;
ConvertTablePGI('DISPO',PanelPGI) ;
{Affaire}
ConvertTablePGI('ACTIVITE',PanelPGI) ;
ConvertTablePGI('AFFAIRE',PanelPGI) ;
ConvertTablePGI('AFFTIERS',PanelPGI) ;
ConvertTablePGI('FACTAFF',PanelPGI) ;
ConvertTablePGI('FONCTION',PanelPGI) ;
ConvertTablePGI('HISTOACTIVITE',PanelPGI) ;
ConvertTablePGI('PROFILGENER',PanelPGI) ;
ConvertTablePGI('RESSOURCE',PanelPGI) ;
Result:=True ;
END ;

Function BasculeEuroPGI ( StDev : String ; PanelPGI : TPanel ) : Boolean ;
BEGIN
OldDev:=StDev ;
if PanelPGI<>Nil then
   BEGIN
   PanelPGI.Visible:=True ;
   PanelPGI.Caption:='Basculement Gestion Commerciale' ;
   PanelPGI.BringToFront ; Application.ProcessMessages ;
   END ;
Result:=BasculeEuroGescom(PanelPGI) ;
if PanelPGI<>Nil then PanelPGI.Visible:=False ;
END ;

end.
