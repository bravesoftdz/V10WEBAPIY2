unit UTofEtatEcartsTransferts;

interface
uses StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,UTOF,EntGc,UTob,hmsgbox,QRS1,
     M3FP,DBTables,HCtrls,HEnt1,FactComm,factutil,CalcOLEGescom,HQry,EdtEtat,voirtob;
{uses   UtilTarif,HDimension,TarifUtil,
    DBTables,

    CalcOle,

ParamSoc,UtilArticle ;}
Type
    TOF_EtatEcartsTransferts = Class(TOF)
    Procedure OnArgument (Argument : String ) ; override ;
    Procedure OnClose ; override ;
    procedure LanceEdition ;
    procedure ChangeCritere ;

    END ;

     var Critere : string;
implementation

Procedure TOF_EtatEcartsTransferts.OnArgument (Argument : String ) ;
var stArgument : string;
begin
inherited ;
stArgument := Argument;
Critere:=uppercase(Trim(ReadTokenSt(stArgument))) ;
if critere = 'TEM'  then Ecran.Caption:=TraduireMemoire('Edition des écarts entre transfert émis et reçu');
if critere = 'TRV' then Ecran.Caption:=TraduireMemoire('Edition des écarts entre transfert à valider et reçu');
updatecaption(Ecran);
if (VH_GC.GCMultiDepots) then
   begin
   SetControlCaption('TGP_DEPOT','Dépôt émetteur');
   SetControlCaption('TGP_DEPOTDEST','Dépôt récepteur');
   end;
end ;

Procedure TOF_EtatEcartsTransferts.OnClose ;
begin
  // Suppression des enregistrements concernant cet utilisateur
  ExecuteSQL('DELETE FROM GCTMPECART WHERE GZT_UTILISATEUR = "'+V_PGI.USer+'"');
  VH_GC.TOBEdt.ClearDetail ;
end ;

procedure TOF_EtatEcartsTransferts.ChangeCritere ;
begin
VH_GC.TOBEdt.ClearDetail ;
end ;

// Procedure de lancement de l'état
procedure TOF_EtatEcartsTransferts.LanceEdition ;
VAR
    stUpdate,stwhere,stTransfertEmis,stTransfertRecu,stInsert,Article,SoucheTEM,SoucheTRE,NaturePieceTEM,NaturePieceTRE,DateTRE : String ;
    refart_sav,Emetteur,Recepteur,DepotEmetteur,DepotRecepteur,Devise,CodeArticle,TypeTransfert,CodeArtPrecedent : String ;
    F : TFQRS1 ;
    QTransfertEmis,QTransfertRecu,QNumLigne : TQuery ;
    TOBTEM,TOBTRE,TOBTEMP,TOBTREArt,TobTempTREArt : TOB ;
    cpt_sav,cptdiff,indexTRE,numprecedent,i_ind1,i_ind2,i_ind3,NumeroTEM,IndicegTEM,NumLigneTEM : integer ;
    NumeroTRE,IndicegTRE,NumLigneTRE,MaxLigne,cpt : integer ;
    QteRecue,QteEmise : double ;
    CleDoc : R_CleDoc ;
    DatePiece,LaDate : TDateTime ;
    Edit : THEdit ;
    Edit2 : THValcombobox ;
    ctrl,ctrlName,signe,refArticle,StTREArt,stWhereNumero : string ;
    choixArticle : boolean;
BEGIN
F := TFQRS1(Ecran);
numprecedent:=0;
cpt:=0;
cptDiff:=-1;
cpt_sav:=0 ;
choixArticle:=false;
TOBTRE:=Nil ; TOBTEMP:=nil ;TOBTREArt:=nil ;
SoucheTEM := '' ; NumeroTEM := 0 ; IndicegTEM := 0 ;

// Suppression des enregistrements concernant cet utilisateur
ExecuteSQL('DELETE FROM GCTMPECART WHERE GZT_UTILISATEUR = "'+V_PGI.USer+'"');

//***************************** CONSTRUCTION DU WHERE ************************
// Le critère DEPOT émetteur
ctrl:='GP_DEPOT';
ctrlName:='DEPOT'; signe:='=' ;
Edit2:=THValComboBox(TFQRS1(F).FindComponent(ctrlName)) ;
if (Edit2 <> nil) and (Edit2.Value <> '') and (Edit2.Value <> TraduireMemoire ('<<Tous>>')) then
   begin
   if stWhere<>'' then stWhere:=stWhere+' AND ' ;
   stWhere:=stWhere+ctrl+signe+'"'+Edit2.Value+'"' ;
   end ;
// Le critère DEPOT récepteur
ctrl:='GP_DEPOTDEST';
ctrlName:='DEPOTDEST'; signe:='=' ;
Edit2:=THValComboBox(TFQRS1(F).FindComponent(ctrlName)) ;
if (Edit2 <> nil) and (Edit2.Value <> '') and (Edit2.Value <> TraduireMemoire ('<<Tous>>')) then
   begin
   if stWhere<>'' then stWhere:=stWhere+' AND ' ;
   stWhere:=stWhere+ctrl+signe+'"'+Edit2.Value+'"' ;
   end ;

// ********************************************************
//Le critère CODE ARTICLE
// pour articles dimensionnés et uniques
ctrl:='GL_ARTICLE';
ctrlName:='REFARTSAISIE'; signe:=' like ' ;
Edit:=THEdit(TFQRS1(F).FindComponent(ctrlName)) ;
if (Edit<>nil) and (Edit.Text<>'') then
   begin
   if stWhere<>'' then stWhere:=stWhere+' AND (' ;
   if stWhere= '' then stWhere:=' (';
   stWhere:=stWhere+ctrl+signe+'"'+Edit.Text+'%"' ;
   choixArticle:=True;
   end ;
// pour articles génériques
ctrl:='GL_CODESDIM';
signe:=' = ' ;
if (Edit<>nil) and (Edit.Text<>'') then
   begin
   if stWhere<>'' then stWhere:=stWhere+' OR ' ;
   stWhere:=stWhere+ctrl+signe+'"'+Edit.Text+'" )' ;
   end ;
// ********************************************************
//Le critère DATEPIECE
ctrl:='GP_DATEPIECE';
ctrlName:='DATEPIECE'; signe:='>=' ;
Edit:=THEdit(TFQRS1(F).FindComponent(ctrlName)) ;
DatePiece:=StrToDate(Edit.Text) ;
if (Edit<>nil) and (Edit.Text<>'') then
   begin
   if stWhere<>'' then stWhere:=stWhere+' AND ' ;
   stWhere:=stWhere+ctrl+signe+'"'+USDateTime(StrToDate(Edit.Text))+'"' ;
   end ;

// Le critère NUMERO et NUMERO_
ctrl:='GP_NUMERO';
ctrlName:='NUMERO'; signe:='>=' ;
Edit:=THEdit(TFQRS1(F).FindComponent(ctrlName)) ;
if (Edit<>nil) and (Edit.Text<>'') then
   begin
   if stWhereNumero<>'' then stWhereNumero:=stWhereNumero+' AND ' ;
   stWhereNumero:=stWhereNumero+ctrl+signe+'"'+Edit.Text+'"' ;
   end ;

ctrl:='GP_NUMERO';
ctrlName:='NUMERO_'; signe:='<=' ;
Edit:=THEdit(TFQRS1(F).FindComponent(ctrlName)) ;
if (Edit<>nil) and (Edit.Text<>'') then
   begin
   if stWhereNumero<>'' then stWhereNumero:=stWhereNumero+' AND ' ;
   stWhereNumero:=stWhereNumero+ctrl+signe+'"'+Edit.Text+'"' ;
   end ;
//****************************************************************************
TypeTransfert:=THEdit(TForm(Ecran).FindComponent('TRANSFERT')).Text ;
stTransfertEmis := 'SELECT gl_naturepieceg,gl_numero,gl_indiceg,gl_souche,gl_numligne,gl_qtefact,'+
'gl_article,gl_codearticle,gl_datepiece,gl_devise,gp_depot,gl_tiers,gp_depotdest,gl_refartsaisie,gl_codesdim'+
' FROM piece'+
' left join ligne on gl_naturepieceg=gp_naturepieceg'+
' and gl_numero=gp_numero and gl_souche=gp_souche'+
' and gl_indiceg=gp_indiceg WHERE (gp_naturepieceg="'+TypeTransfert+'")';
if stWhereNumero<>'' then stTransfertEmis :=stTransfertEmis+' and '+stWhereNumero;
stTransfertEmis :=stTransfertEmis+' and '+stWhere+
' and gp_indiceg=0 '+
' order by gl_numero,gl_indiceg,gl_souche,gl_numligne' ;

// Si la TOB n'existe pas ou qu'il y a aucun enregistrement alors
// je la crée et la remplie
TOBTEM:=VH_GC.TOBEdt.FindFirst(['_TEM'],['TEM'],false);
if (TOBTEM=nil) or (TOBTEM.Detail.count=0) then
   begin
   TOBTEM:=TOB.Create('',nil,-1) ;
   TOBTEM.AddChampSup('_TEM',False) ;
   TOBTEM.PutValue('_TEM','TEM') ;
   QTransfertEmis:=OpenSQL(stTransfertEmis,True) ;
   TOBTEM.LoadDetailDB('LIGNE','','',QTransfertEmis,False) ;
   if TOBTEM<>nil then TOBTEM.Changeparent(VH_GC.TOBEdt,-1) ;
   Ferme(QTransfertEmis) ;
   end ;

// Si on choisi un codearticle comme critère, je créé une nouvelle TOB qui récupère
// tous les transferts reçus de cet article pour effectué un traitement particulier
// si cet article est présent dans le TRE mais pas dans le TEM ou TRV
if choixArticle then
begin
StTREArt := 'SELECT gl_naturepieceg,gl_numero,gl_indiceg,gl_souche,gl_numligne,gl_qtefact,'+
'gl_article,gl_codearticle,gl_datepiece,gl_devise,gp_depot,gl_tiers,gp_depotdest,gl_refartsaisie,gl_codesdim'+
' FROM piece'+
' left join ligne on gl_naturepieceg=gp_naturepieceg'+
' and gl_numero=gp_numero and gl_souche=gp_souche'+
' and gl_indiceg=gp_indiceg WHERE (gp_naturepieceg="TRE")';
if stWhereNumero<>'' then StTREArt :=StTREArt+' AND '+stWhereNumero;
StTREArt :=StTREArt+' and '+stWhere+
' and  and gp_indiceg=0 '+
' order by gl_numero,gl_indiceg,gl_souche,gl_numligne' ;
if existeSQL(StTREArt) then  // si la requête ramène qqchose je crée la TOB correspondante
   begin
   TOBTREArt:=VH_GC.TOBEdt.FindFirst(['_TREArt'],['TREArt'],false);
   if (TOBTREArt=nil) or (TOBTREArt.Detail.count=0) then
      begin
      TOBTREArt:=TOB.Create('',nil,-1) ;
      TOBTREArt.AddChampSup('_TREArt',False) ;
      TOBTREArt.PutValue('_TREArt','TREArt') ;
      QTransfertEmis:=OpenSQL(StTREArt,True) ;
      TOBTREArt.LoadDetailDB('','','',QTransfertEmis,False) ;
      if TOBTREArt<>nil then TOBTREArt.Changeparent(VH_GC.TOBEdt,-1) ;
      Ferme(QTransfertEmis) ;
      end ;
   end;
end;
// Je parcours tous les transferts émis
for i_ind1:=0 to TOBTEM.Detail.count-1 do
   begin
   // Si on est toujours dans le même transfert, il faut vérifier que dans le transfert reçu,
   // il ne reste pas des lignes ayant le même code article pour que dans l'édition tous les
   // articles ayant le même code article soit les uns en dessous des autres
   if numprecedent=TOBTEM.Detail[i_ind1].getvalue('GL_NUMERO') then
      begin
      // pour avoir le code de l'article il faut si c'est un article générique le récupérer par GL_CODESDIM
      // sinon on le récupère par GL_CODEARTICLE
      if TOBTEM.Detail[i_ind1].GetValue('GL_CODEARTICLE')<>'' then refArticle:='GL_CODEARTICLE' else refArticle:='GL_CODESDIM';
      if CodeArtPrecedent<>TOBTEM.Detail[i_ind1].GetValue(refArticle) then
         begin
         TOBTEMP:= TOBTRE.FindFirst(['GL_NATUREPIECEG','GL_SOUCHE','GL_NUMERO','GL_INDICEG',refArticle],['TRE',SoucheTEM,NumeroTEM,IndicegTEM,CodeArtPrecedent],True);
         while TOBTEMP<>nil do
            begin
            NaturePieceTEM:=TypeTransfert;
            NaturePieceTRE:=TOBTEMP.GetValue('GL_NATUREPIECEG');
            IF Not VarIsNull(TOBTEMP.GetValue('GL_ARTICLE')) THEN Article:=TOBTEMP.GetValue('GL_ARTICLE')
            else Article:='';
            if Not VarIsNull(TOBTEMP.GetValue('GL_CODEARTICLE')) then
            begin
            if TOBTEMP.GetValue('GL_CODEARTICLE')<>'' then refArticle:='GL_CODEARTICLE' else refArticle:='GL_CODESDIM';
            end
            else refArticle:='GL_CODESDIM';
            CodeArticle:=TOBTEMP.GetValue(refArticle);
            DateTRE:=TOBTEMP.GetValue('GL_DATEPIECE');
            NumeroTRE:=TOBTEMP.GetValue('GL_NUMERO');
            SoucheTRE:=TOBTEMP.GetValue('GL_SOUCHE');
            IndicegTRE:=TOBTEMP.GetValue('GL_INDICEG');
            NumLigneTRE:=TOBTEMP.GetValue('GL_NUMLIGNE');
            DepotEmetteur:=TOBTEMP.GetValue('GP_DEPOT');
            DepotRecepteur:=TOBTEMP.GetValue('GP_DEPOTDEST');
            Devise:=TOBTEMP.GetValue('GL_DEVISE');
            QteRecue:=TOBTEMP.getvalue('GL_QTEFACT');

            // Création des chaine de caractères "Emetteur" et "Recepteur" pour les insérer
            Emetteur:=NaturePieceTEM+';'+SoucheTRE+';'+IntToStr(NumeroTRE)+';'
                      +IntToStr(IndicegTRE);
            Recepteur:=NaturePieceTRE+';'+SoucheTRE+';'+IntToStr(NumeroTRE)+';'
                      +IntToStr(IndicegTRE);
            // Création du Insert dans la table temporaire
            // increment du compteur
            cpt:=cpt+1;
            if article='' then
               begin
               if cptDiff<>-1 then
                  begin
                  stUpdate:='Update GCTMPECART set GZT_QTEEMISE='+IntToStr(cptDiff)
                  +' Where GZT_COMPTEUR='+IntToStr(cpt_sav);
                  ExecuteSQL(stUpdate);
                  end;
               cptDiff:=0;
               cpt_sav:=cpt;
               refart_sav:=CodeArticle;
               end
            else if CodeArticle=refart_sav then  inc(cptDiff);
            stInsert:='Insert into GCTMPECART (GZT_UTILISATEUR,GZT_COMPTEUR,GZT_EMETTEUR,GZT_RECEPTEUR,GZT_NUMLIGRE,GZT_ARTICLE,'
                      +'GZT_CODEARTICLE,GZT_DEPOTEM,GZT_DEPOTRE,GZT_DEVISE,GZT_DATEEM,GZT_DATERE,GZT_QTERECUE)'
                      +'values ("'+V_PGI.USer+'",'+IntToStr(cpt)+',"'+Emetteur+'","'+Recepteur+'",'+IntToStr(NumLigneTRE)+',"'
                      +Article+'","'+CodeArticle+'","'+DepotEmetteur+'","'
                      +DepotRecepteur+'","'+Devise+'","'+DateToStr(DatePiece)+'","'+DateTRE+'",'
                      +FloatToStr(QteRecue)+')';
            ExecuteSQL(stInsert);
            TOBTEMP.free;
            TOBTEMP:= TOBTRE.FindNext(['GL_NATUREPIECEG','GL_SOUCHE','GL_NUMERO','GL_INDICEG',refArticle],['TRE',SoucheTEM,NumeroTEM,IndicegTEM,CodeArtPrecedent],True);
            end;
         stUpdate:='Update GCTMPECART set GZT_QTEEMISE='+IntToStr(cptDiff)
                   +' Where GZT_COMPTEUR='+IntToStr(cpt_sav);
         ExecuteSQL(stUpdate);
         // Réinitialisation du code article en cours
         CodeArtPrecedent:=TOBTEM.Detail[i_ind1].GetValue('GL_REFARTSAISIE');
         end;
      end;

   // Je ne traite l'ALF que si le numero de l'ALF est différent du précédent
   if numprecedent<>TOBTEM.Detail[i_ind1].getvalue('GL_NUMERO') then
      begin
      // Réinitialisation du code article en cours
      if TOBTEM.Detail[i_ind1].GetValue('GL_CODEARTICLE')<>'' then refArticle:='GL_CODEARTICLE' else refArticle:='GL_CODESDIM';
      CodeArtPrecedent:=TOBTEM.Detail[i_ind1].GetValue(refArticle);
      // Ceci correspond au cas ou l'on a plus de ligne dans le transfert reçu que dans le transfert émi
      // alors on traite toutes les lignes du transfert reçu 
      if (TOBTRE<>nil) and (TOBTRE.Detail.count<>0) then
         begin
         for i_ind3:=0 to TOBTRE.Detail.count-1 do
            begin
            NaturePieceTEM:=TypeTransfert;
            NaturePieceTRE:=TOBTRE.Detail[i_ind3].GetValue('GL_NATUREPIECEG');
            IF Not VarIsNull(TOBTRE.Detail[i_ind3].GetValue('GL_ARTICLE')) THEN Article:=TOBTRE.Detail[i_ind3].GetValue('GL_ARTICLE')
            else Article:='';
            if Not VarIsNull(TOBTRE.Detail[i_ind3].GetValue('GL_CODEARTICLE')) then
            begin
            if TOBTRE.Detail[i_ind3].GetValue('GL_CODEARTICLE')<>'' then refArticle:='GL_CODEARTICLE' else refArticle:='GL_CODESDIM';
            end
            else refArticle:='GL_CODESDIM';
            CodeArticle:=TOBTRE.Detail[i_ind3].GetValue(refArticle);
            DateTRE:=TOBTRE.Detail[i_ind3].GetValue('GL_DATEPIECE');
            NumeroTRE:=TOBTRE.Detail[i_ind3].GetValue('GL_NUMERO');
            SoucheTRE:=TOBTRE.Detail[i_ind3].GetValue('GL_SOUCHE');
            IndicegTRE:=TOBTRE.Detail[i_ind3].GetValue('GL_INDICEG');
            NumLigneTRE:=TOBTRE.Detail[i_ind3].GetValue('GL_NUMLIGNE');
            DepotEmetteur:=TOBTRE.Detail[i_ind3].GetValue('GP_DEPOT');
            DepotRecepteur:=TOBTRE.Detail[i_ind3].GetValue('GP_DEPOTDEST');
            Devise:=TOBTRE.Detail[i_ind3].GetValue('GL_DEVISE');
            QteRecue:=TOBTRE.Detail[i_ind3].getvalue('GL_QTEFACT');
            // Création des chaine de caractères "Emetteur" et "Recepteur" pour les insérer
            Emetteur:=NaturePieceTEM+';'+SoucheTRE+';'+IntToStr(NumeroTRE)+';'
                +IntToStr(IndicegTRE);
            Recepteur:=NaturePieceTRE+';'+SoucheTRE+';'+IntToStr(NumeroTRE)+';'
                +IntToStr(IndicegTRE);
            // Création du Insert dans la table temporaire
            // increment du compteur
            cpt:=cpt+1;
            if article='' then
               begin
               if cptDiff<>-1 then
                  begin
                  stUpdate:='Update GCTMPECART set GZT_QTEEMISE='+IntToStr(cptDiff)
                               +' Where GZT_COMPTEUR='+IntToStr(cpt_sav);
                  ExecuteSQL(stUpdate);
                  end;
               cptDiff:=0;
               cpt_sav:=cpt;
               refart_sav:=CodeArticle;
               end
            else if CodeArticle=refart_sav then  inc(cptDiff);
            stInsert:='Insert into GCTMPECART (GZT_UTILISATEUR,GZT_COMPTEUR,GZT_EMETTEUR,GZT_RECEPTEUR,GZT_NUMLIGRE,GZT_ARTICLE,'
                      +'GZT_CODEARTICLE,GZT_DEPOTEM,GZT_DEPOTRE,GZT_DEVISE,GZT_DATEEM,GZT_DATERE,GZT_QTERECUE)'
                      +'values ("'+V_PGI.USer+'",'+IntToStr(cpt)+',"'+Emetteur+'","'+Recepteur+'",'
                      +IntToStr(NumLigneTRE)+',"'+Article+'","'+CodeArticle+'","'+DepotEmetteur+'","'
                      +DepotRecepteur+'","'+Devise+'","'+DateToStr(DatePiece)+'","'+DateTRE+'",'
                      +FloatToStr(QteRecue)+')';
            ExecuteSQL(stInsert);
            end;
            stUpdate:='Update GCTMPECART set GZT_QTEEMISE='+IntToStr(cptDiff)
                         +' Where GZT_COMPTEUR='+IntToStr(cpt_sav);
            ExecuteSQL(stUpdate);
         end;
      // Il faut détruire la TOB des transferts reçus si elle n'est pas vide,
      // on peut détruire tous les enregistrements car on les a traité avant.
      if (TOBTRE<>nil) then
      begin
      TOBTRE.free;
      end;
      // Récupération du transfert émis correspondant, des valeurs qui ne change que si le numéro change
      NaturePieceTEM:=TOBTEM.Detail[i_ind1].GetValue('GL_NATUREPIECEG');
      DatePiece:=TOBTEM.Detail[i_ind1].GetValue('GL_DATEPIECE');
      NumeroTEM:=TOBTEM.Detail[i_ind1].GetValue('GL_NUMERO');
      SoucheTEM:=TOBTEM.Detail[i_ind1].GetValue('GL_SOUCHE');
      IndicegTEM:=TOBTEM.Detail[i_ind1].GetValue('GL_INDICEG');
      DepotEmetteur:=TOBTEM.Detail[i_ind1].GetValue('GP_DEPOT');
      DepotRecepteur:=TOBTEM.Detail[i_ind1].GetValue('GP_DEPOTDEST');
      Devise:=TOBTEM.Detail[i_ind1].GetValue('GL_DEVISE');
      // récupération du numéro en cours
      numprecedent:=TOBTEM.Detail[i_ind1].getvalue('GL_NUMERO');
      // Création de la tob des receptions
      stTransfertRecu:= 'SELECT gl_naturepieceg,gl_numero,gl_indiceg,gl_souche,gl_numligne,gl_qtefact,'+
      'gl_article,gl_codearticle,gl_datepiece,gl_devise,gl_pieceprecedente,gl_refartsaisie,gl_codesdim,'+
      'gp_depot,gp_depotdest'+
      ' FROM piece'+
      ' left join ligne on gl_naturepieceg=gp_naturepieceg'+
      ' and gl_numero=gp_numero and gl_souche=gp_souche'+
      ' and gl_indiceg=gp_indiceg '+
      ' WHERE gp_naturepieceg="TRE" and '+stWhere+' and gp_numero ='+IntToStr(numprecedent)+
      ' and gp_indiceg=0 ';
      TOBTRE:=TOB.Create('',nil,-1) ;
      QTransfertRecu:=OpenSQL(stTransfertRecu,True) ;
      if not QTransfertRecu.EOF then TOBTRE.LoadDetailDB('', '', '', QTransfertRecu, FAlse);
      Ferme (QTransfertRecu) ;
      end;  // fin traitement qd numero transfert différent de numéro precedent

   stInsert:='';

   // Récupération du transfert émis correspondant, des autres valeurs
   Article:=TOBTEM.Detail[i_ind1].GetValue('GL_ARTICLE');
   if TOBTEM.Detail[i_ind1].GetValue('GL_CODEARTICLE')<>''then refArticle:='GL_CODEARTICLE' else refArticle:='GL_CODESDIM';
   CodeArticle:=TOBTEM.Detail[i_ind1].GetValue(refArticle);
   NumLigneTEM:=TOBTEM.Detail[i_ind1].GetValue('GL_NUMLIGNE');
   QteEmise:=TOBTEM.Detail[i_ind1].GetValue('GL_QTEFACT');


   // Si il on a des transfert reçus
   if (TOBTRE<>nil) and (TOBTRE.Detail.count<>0) then
      begin
      // on regarde si la ligne du TEM en cours existe dans les transfert reçus
      if Article='' then
         begin
         TOBTEMP:= TOBTRE.FindFirst(['GL_NATUREPIECEG','GL_SOUCHE','GL_NUMERO','GL_INDICEG',refArticle,'GL_ARTICLE'],['TRE',SoucheTEM,NumeroTEM,IndicegTEM,CodeArticle,null],True);
         if TOBTEMP=nil then TOBTEMP:= TOBTRE.FindFirst(['GL_NATUREPIECEG','GL_SOUCHE','GL_NUMERO','GL_INDICEG',refArticle,'GL_ARTICLE'],['TRE',SoucheTEM,NumeroTEM,IndicegTEM,CodeArticle,article],True);
         end
      else TOBTEMP:= TOBTRE.FindFirst(['GL_NATUREPIECEG','GL_SOUCHE','GL_NUMERO','GL_INDICEG',refArticle,'GL_ARTICLE'],['TRE',SoucheTEM,NumeroTEM,IndicegTEM,CodeArticle,Article],True);
      end;
   if (TOBTEMP<>nil) then
      begin
      // Si la ligne existe je récupère son index
      i_ind2:=TOBTEMP.GetIndex ;
      TOBTEMP:=nil;
      // Récupération du transfert reçu correspondant
      NaturePieceTRE:=TOBTRE.Detail[i_ind2].GetValue('GL_NATUREPIECEG');
      NumeroTRE:=TOBTRE.Detail[i_ind2].GetValue('GL_NUMERO');
      SoucheTRE:=TOBTRE.Detail[i_ind2].GetValue('GL_SOUCHE');
      IndicegTRE:=TOBTRE.Detail[i_ind2].GetValue('GL_INDICEG');
      NumLigneTRE:=TOBTRE.Detail[i_ind2].GetValue('GL_NUMLIGNE');
      QteRecue:=TOBTRE.Detail[i_ind2].GetValue('GL_QTEFACT');

      // Création des chaine de caractères "Emetteur" et "Recepteur" pour les insérer
      Emetteur:=NaturePieceTEM+';'+SoucheTEM+';'+IntToStr(NumeroTEM)+';'+IntToStr(IndicegTEM);
      Recepteur:=NaturePieceTRE+';'+SoucheTRE+';'+IntToStr(NumeroTRE)+';'+IntToStr(IndicegTRE) ;
      // Création du Insert dans la table temporaire
      DateTRE:= TOBTRE.Detail[i_ind2].getvalue('GL_DATEPIECE');
      // quand article ='' il faut vérifier si il a des quantités différentes au niveau des articles dimensionnés
      // en effet regarder juste sur la différence entre qté emise et recue ne suffit pas car on peut avoir une
      // même qté totale sur une réference mais une répartition différente au niveau des articles dimensionné
      if (article='') or (QteEmise<>QteRecue) then
         begin
         // increment du compteur
         cpt:=cpt+1;
         if article='' then
            begin
            if cptDiff<>-1 then
               begin
               stUpdate:='Update GCTMPECART set GZT_QTEEMISE='+IntToStr(cptDiff)
                            +' Where GZT_COMPTEUR='+IntToStr(cpt_sav);
               ExecuteSQL(stUpdate);
               end;
            cptDiff:=0;
            cpt_sav:=cpt;
            refart_sav:=CodeArticle;
            end
         else if CodeArticle=refart_sav then  inc(cptDiff);
         stInsert:='Insert into GCTMPECART values ("'+V_PGI.USer+'",'+IntToStr(cpt)+',"'+Emetteur+'",'+IntToStr(NumLigneTEM)+',"'
                   +Recepteur+'",'+IntToStr(NumLigneTRE)+',"'
                   +Article+'","'+CodeArticle+'","'+DepotEmetteur+'","'+DepotRecepteur+'","'
                   +'","'+Devise+'","'+DateToStr(DatePiece)+'","'+DateTRE+'",'
                   +FloatToStr(QteEmise)+','+FloatToStr(QteRecue)+')';
         ExecuteSQL(stInsert);
         end;
      // je regarde si il existe dans TOBTREArt si oui je le supprime car je viens de le traiter
      if choixArticle then
         begin
         TobTempTREArt:=TOBTREArt.FindFirst(['GL_NATUREPIECEG','GL_SOUCHE','GL_NUMERO','GL_INDICEG','GL_NUMLIGNE'],['TRE',SoucheTRE,IntToStr(NumeroTRE),IntToStr(IndicegTRE),IntToStr(NumLigneTRE)],True);
         if TobTempTREArt<>nil then TobTempTREArt.free;
         end;
      TOBTRE.Detail[i_ind2].free;
      if (i_ind1=TOBTEM.Detail.count-1) and (TOBTRE.Detail.count=0) then
         begin
         stUpdate:='Update GCTMPECART set GZT_QTEEMISE='+IntToStr(cptDiff)
                      +' Where GZT_COMPTEUR='+IntToStr(cpt_sav);
         ExecuteSQL(stUpdate);
         end;
      end
   else
      begin
      // Création des chaine de caractères "Emetteur"
      Emetteur:=NaturePieceTEM+';'+SoucheTEM+';'+IntToStr(NumeroTEM)+';'
      +IntToStr(IndicegTEM);
      // increment du compteur
      cpt:=cpt+1;
      if article='' then
         begin
         if cptDiff<>-1 then
            begin
            stUpdate:='Update GCTMPECART set GZT_QTEEMISE='+IntToStr(cptDiff)
                         +' Where GZT_COMPTEUR='+IntToStr(cpt_sav);
            ExecuteSQL(stUpdate);
            end;
         cptDiff:=0;
         cpt_sav:=cpt;
         refart_sav:=CodeArticle;
         end
      else if CodeArticle=refart_sav then  inc(cptDiff);
      // Création du Insert dans la table temporaire
      stInsert:='Insert into GCTMPECART (GZT_UTILISATEUR,GZT_COMPTEUR,GZT_EMETTEUR,GZT_NUMLIGEM,GZT_ARTICLE,'
                +'GZT_CODEARTICLE,GZT_DEPOTEM,GZT_DEPOTRE,GZT_DEVISE,GZT_DATEEM,GZT_QTEEMISE)'
                +' values ("'+V_PGI.USer+'",'+IntToStr(cpt)+',"'+Emetteur+'",'+IntToStr(NumLigneTEM)+',"'
                +Article+'","'+CodeArticle+'","'+DepotEmetteur+'","'+DepotRecepteur+'","'
                +Devise+'","'+DateToStr(DatePiece)+'",'+FloatToStr(QteEmise)+')';
      ExecuteSQL(stInsert);
      if i_ind1=TOBTEM.Detail.count-1 then
         begin
         stUpdate:='Update GCTMPECART set GZT_QTEEMISE='+IntToStr(cptDiff)
                      +' Where GZT_COMPTEUR='+IntToStr(cpt_sav);
         ExecuteSQL(stUpdate);
         end;
      end;
   end; // fin pour les transferts émis
// je traite les transferts reçus restant
if (TOBTRE<>nil) and (TOBTRE.Detail.count<>0) then
   begin
   for i_ind3:=0 to TOBTRE.Detail.count-1 do
      begin
      NaturePieceTEM:=TypeTransfert;
      NaturePieceTRE:=TOBTRE.Detail[i_ind3].GetValue('GL_NATUREPIECEG');
      IF Not VarIsNull(TOBTRE.Detail[i_ind3].GetValue('GL_ARTICLE')) then Article:=TOBTRE.Detail[i_ind3].GetValue('GL_ARTICLE')
      else Article:='';
      IF Not VarIsNull(TOBTRE.Detail[i_ind3].GetValue('GL_CODEARTICLE')) then
      begin
      if TOBTRE.Detail[i_ind3].GetValue('GL_CODEARTICLE') <>'' then refArticle:='GL_CODEARTICLE' else refArticle:='GL_CODESDIM';
      end
      else refArticle:='GL_CODESDIM';
      CodeArticle:=TOBTRE.Detail[i_ind3].GetValue(refArticle);
      DateTRE:=TOBTRE.Detail[i_ind3].GetValue('GL_DATEPIECE');
      NumeroTRE:=TOBTRE.Detail[i_ind3].GetValue('GL_NUMERO');
      SoucheTRE:=TOBTRE.Detail[i_ind3].GetValue('GL_SOUCHE');
      IndicegTRE:=TOBTRE.Detail[i_ind3].GetValue('GL_INDICEG');
      NumLigneTRE:=TOBTRE.Detail[i_ind3].GetValue('GL_NUMLIGNE');
      DepotEmetteur:=TOBTRE.Detail[i_ind3].GetValue('GP_DEPOT');
      DepotRecepteur:=TOBTRE.Detail[i_ind3].GetValue('GP_DEPOTDEST');
      Devise:=TOBTRE.Detail[i_ind3].GetValue('GL_DEVISE');
      QteRecue:=TOBTRE.Detail[i_ind3].getvalue('GL_QTEFACT');
      // Création des chaine de caractères "Emetteur" et "Recepteur" pour les insérer
      Emetteur:=NaturePieceTEM+';'+SoucheTRE+';'+IntToStr(NumeroTRE)+';'
                +IntToStr(IndicegTRE);
      Recepteur:=NaturePieceTRE+';'+SoucheTRE+';'+IntToStr(NumeroTRE)+';'
                +IntToStr(IndicegTRE);
      // Création du Insert dans la table temporaire
      // increment du compteur
      cpt:=cpt+1;
      if article='' then
         begin
         if cptDiff<>-1 then
            begin
            stUpdate:='Update GCTMPECART set GZT_QTEEMISE='+IntToStr(cptDiff)
                         +' Where GZT_COMPTEUR='+IntToStr(cpt_sav);
            ExecuteSQL(stUpdate);
            end;
         cptDiff:=0;
         cpt_sav:=cpt;
         refart_sav:=CodeArticle;
         end
      else if CodeArticle=refart_sav then  inc(cptDiff);
      stInsert:='Insert into GCTMPECART (GZT_UTILISATEUR,GZT_COMPTEUR,GZT_EMETTEUR,GZT_RECEPTEUR,GZT_NUMLIGRE,GZT_ARTICLE,'
                +'GZT_CODEARTICLE,GZT_DEPOTEM,GZT_DEPOTRE,GZT_DEVISE,GZT_DATEEM,GZT_DATERE,GZT_QTERECUE)'
                +'values ("'+V_PGI.USer+'",'+IntToStr(cpt)+',"'+Emetteur+'","'+Recepteur+'",'
                +IntToStr(NumLigneTRE)+',"'+Article+'","'+CodeArticle+'","'+DepotEmetteur+'","'
                +DepotRecepteur+'","'+Devise+'","'+DatetoStr(DatePiece)+'","'+DateTRE+'",'
                +FloatToStr(QteRecue)+')';
      ExecuteSQL(stInsert);
      // je regarde si il existe dans TOBTREArt si oui je le supprime car je viens de le traiter
      if choixArticle then
         begin
         TobTempTREArt:=TOBTREArt.FindFirst(['GL_NATUREPIECEG','GL_SOUCHE','GL_NUMERO','GL_INDICEG','GL_NUMLIGNE'],['TRE',SoucheTRE,IntToStr(NumeroTRE),IntToStr(IndicegTRE),IntToStr(NumLigneTRE)],True);
         if TobTempTREArt<>nil then TobTempTREArt.free;
         end;
      end;
      stUpdate:='Update GCTMPECART set GZT_QTEEMISE='+IntToStr(cptDiff)
                   +' Where GZT_COMPTEUR='+IntToStr(cpt_sav);
      ExecuteSQL(stUpdate);
   end; // fin transfert reçu restant

// je traite TOBTREArt si elle n'est pas vide
if (TOBTREArt<>nil) then
   begin
   // Si j'ai selectionné un article et que la TOBTREArt n'est pas vide alors je traite ces enregistrements
   for i_ind3:=0 to TOBTREArt.Detail.count-1 do
      begin
      NaturePieceTEM:=TypeTransfert;
      NaturePieceTRE:=TOBTREArt.Detail[i_ind3].GetValue('GL_NATUREPIECEG');
      IF Not VarIsNull(TOBTREArt.Detail[i_ind3].GetValue('GL_ARTICLE')) then Article:=TOBTREArt.Detail[i_ind3].GetValue('GL_ARTICLE')
      else Article:='';
      IF Not VarIsNull(TOBTREArt.Detail[i_ind3].GetValue('GL_CODEARTICLE')) then
      begin
      if TOBTREArt.Detail[i_ind3].GetValue('GL_CODEARTICLE') <>'' then refArticle:='GL_CODEARTICLE' else refArticle:='GL_CODESDIM';
      end
      else refArticle:='GL_CODESDIM';
      CodeArticle:=TOBTREArt.Detail[i_ind3].GetValue(refArticle);
      DateTRE:=TOBTREArt.Detail[i_ind3].GetValue('GL_DATEPIECE');
      NumeroTRE:=TOBTREArt.Detail[i_ind3].GetValue('GL_NUMERO');
      SoucheTRE:=TOBTREArt.Detail[i_ind3].GetValue('GL_SOUCHE');
      IndicegTRE:=TOBTREArt.Detail[i_ind3].GetValue('GL_INDICEG');
      NumLigneTRE:=TOBTREArt.Detail[i_ind3].GetValue('GL_NUMLIGNE');
      DepotEmetteur:=TOBTREArt.Detail[i_ind3].GetValue('GP_DEPOT');
      DepotRecepteur:=TOBTREArt.Detail[i_ind3].GetValue('GP_DEPOTDEST');
      Devise:=TOBTREArt.Detail[i_ind3].GetValue('GL_DEVISE');
      QteRecue:=TOBTREArt.Detail[i_ind3].getvalue('GL_QTEFACT');
      // Rechereche de la Date du transfert émi ou à valider
      stTREArt:='Select gp_datepiece from piece where gp_naturepieceg="'+TypeTransfert
                +'" and gp_souche="'+SoucheTRE+'" and gp_numero='+IntToStr(NumeroTRE)
                +' and gp_indiceg='+IntToStr(IndicegTRE);
      QTransfertEmis:=OpenSQL(stTREArt,True);
      DatePiece:=QTransfertEmis.FindField('GP_DATEPIECE').AsDateTime;
      Ferme(QTransfertEmis);
      // Création des chaine de caractères "Emetteur" et "Recepteur" pour les insérer
      Emetteur:=NaturePieceTEM+';'+SoucheTRE+';'+IntToStr(NumeroTRE)+';'
                +IntToStr(IndicegTRE);
      Recepteur:=NaturePieceTRE+';'+SoucheTRE+';'+IntToStr(NumeroTRE)+';'
                +IntToStr(IndicegTRE);
      // Création du Insert dans la table temporaire
      // increment du compteur
      cpt:=cpt+1;
      if article='' then
         begin
         if cptDiff<>-1 then
            begin
            stUpdate:='Update GCTMPECART set GZT_QTEEMISE='+IntToStr(cptDiff)
                         +' Where GZT_COMPTEUR='+IntToStr(cpt_sav);
            ExecuteSQL(stUpdate);
            end;
         cptDiff:=0;
         cpt_sav:=cpt;
         refart_sav:=CodeArticle;
         end
      else if CodeArticle=refart_sav then  inc(cptDiff);
      stInsert:='Insert into GCTMPECART (GZT_UTILISATEUR,GZT_COMPTEUR,GZT_EMETTEUR,GZT_RECEPTEUR,GZT_NUMLIGRE,GZT_ARTICLE,'
                +'GZT_CODEARTICLE,GZT_DEPOTEM,GZT_DEPOTRE,GZT_DEVISE,GZT_DATEEM,GZT_DATERE,GZT_QTERECUE)'
                +'values ("'+V_PGI.USer+'",'+IntToStr(cpt)+',"'+Emetteur+'","'+Recepteur+'",'
                +IntToStr(NumLigneTRE)+',"'+Article+'","'+CodeArticle+'","'+DepotEmetteur+'","'
                +DepotRecepteur+'","'+Devise+'","'+DatetoStr(DatePiece)+'","'+DateTRE+'",'
                +FloatToStr(QteRecue)+')';
      ExecuteSQL(stInsert);
      end;
      stUpdate:='Update GCTMPECART set GZT_QTEEMISE='+IntToStr(cptDiff)
                   +' Where GZT_COMPTEUR='+IntToStr(cpt_sav);
      ExecuteSQL(stUpdate);
   end; // fin TOBTREArt

// je supprime tous les enregistrements dont l'article = '' et que la qteemise=0
// ceci afin d'avoir dans la table que les enregistrements à éditer pour résoudre les
// problème d'édition de l'entête de transfert
ExecuteSQL('Delete GCTMPECART where GZT_ARTICLE="" AND GZT_QTEEMISE=0');

if TOBTRE<>nil then TOBTRE.free;
//if TOBTEM<>nil then TOBTEM.free;
// Clause where afin d'être sûre de récupérer ses propres enregistrements
THEdit(TForm(Ecran).FindComponent('GZT_UTILISATEUR')).Text:= V_PGI.USer;
END ;

// Procedure AGL pour lancer le traitement pour l'édition des écarts
procedure AGLOnClickLanceEditionTransfert( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFQRS1) then TOTOF:=TFQRS1(F).LaTOF
  else exit;
  if (TOTOF is TOF_EtatEcartsTransferts) then TOF_EtatEcartsTransferts(TOTOF).LanceEdition ;
end;
// Procedure AGL pour vider la TOB des transferts emis quand il y a une modification de critères
procedure AGLOnChangeCritere( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFQRS1) then TOTOF:=TFQRS1(F).LaTOF
  else exit;
  if (TOTOF is TOF_EtatEcartsTransferts) then TOF_EtatEcartsTransferts(TOTOF).ChangeCritere ;
end;

Initialization
registerclasses([TOF_EtatEcartsTransferts]);
RegisterAglProc('OnChangeCritere',TRUE,1,AGLOnChangeCritere);
RegisterAglProc('OnClickLanceEditionTransfert',TRUE,1,AGLOnClickLanceEditionTransfert);
end.
