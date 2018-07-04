{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 20/04/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFSBPARBREPREV ()
Mots clefs ... : TOF;QUFSBPARBREPREV
*****************************************************************}
Unit QUFSBPARBREPREV_TOF ;

Interface


Uses Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}fe_main,
{$ELSE}
     MainEagl,   
     utob,      
{$ENDIF}
     sysutils,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,UGraph ;

Type
  TOF_QUFSBPARBREPREV = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure LineEnter(Sender: TObject);
    procedure DblClick(Sender: TObject);
    procedure Laftershow;
   private
    codeSession,structure,VuePgiDelai,VueCoeff,VuePrev:hString;
    chpSql: hString;
    nbTabAffiche:integer;
    nivMax,nivMaxSession,nbValeurTabD:integer;
    TotalCA1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6,TotalQte:double;
    QTBP:TQTob;
    ligne,noeud,coeffC,coeffRet:integer;
    okCoeff,OkPrev,OkInitNivTaille:boolean;
    TabCodeAxe,TabNumValAxe:array [1..11] of hString;
    TabLibCol:array of hString;
    TabD,TabChp : array of String;
    procedure killTob;
    procedure creerTob ;
    procedure chargeTob ;
    procedure ChgColonne() ;
    procedure ClickModif(Sender: TObject);
    procedure recalculTV();
    procedure MiseAJourDonneeBP(const retour,noeudcourant:hString);
    function DonneFiltre:hString;
    procedure ClickDetailClient(Sender: TObject);
    procedure CodeSqlMiseAJourDonnee(RI:integer;retour:hString;
              var codeSql,codeM:hString);
  end ;


Implementation

uses Stat,UTobView,HTB97,menus,
     Uutil,BPFctSession,BPFctArbre,BPBasic,BPUtil,SynScriptBP;

procedure TOF_QUFSBPARBREPREV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFSBPARBREPREV.OnDelete ;
begin
  Inherited ;
end ;


procedure TOF_QUFSBPARBREPREV.OnUpdate ;
begin
  Inherited ;
end;


procedure TOF_QUFSBPARBREPREV.creerTob ;
var RI:integer;
    ChpVirgule,nbVal,codeAxe:hString;
    NbChp,i,j:integer;
    ChpS,ChpColonneD,codeFiltre,codeType:hString;
    TabValAff,TabLibValAff:array [0..7] of hString;
begin
  okCoeff:=false;
  OkPrev:=false;
  Case ContextBp of
   0,1,2 : nbTabAffiche:=1;
   3 : begin
         LibValAff(codeSession,TabLibValAff);
         nbTabAffiche := StrToInt(TabLibValAff[0]);
       end;
  end;

  //cherche le nombre de niveau definit
  nivMax:=ChercheNivMax(codeSession);
  nivMaxSession:=ChercheNivMaxSession(codeSession);
  OkInitNivTaille:=(SessionCalculParTaille(codeSession) or (SessionEclateeParTaille(codeSession)));

  //cherche valeurs totales
  TotalCA1:=ChercheTotalArbre(codeSession,'CA1');
  TotalCA2:=ChercheTotalArbre(codeSession,'CA2');
  TotalCA3:=ChercheTotalArbre(codeSession,'CA3');
  TotalCA4:=ChercheTotalArbre(codeSession,'CA4');
  TotalCA5:=ChercheTotalArbre(codeSession,'CA5');
  TotalCA6:=ChercheTotalArbre(codeSession,'CA6');
  TotalQte:=ChercheTotalArbre(codeSession,'QTE');

  structure:='';
  if BPOkOrli
   then
    //-----------------> ORLI
    begin
     structure:=ChercheSessionStructure(codeSession);
     ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);
    end
    //ORLI <-----------------
   else ChercheCodeAxeSession(codeSession,TabCodeAxe);

  if not BPOkOrli
   then
    begin
     if SessionDelai(codeSession)
      then
       begin
        nivMax:=nivMax+1;
        TabCodeAxe[NivMax+1]:='DELAI';
       end;

     Case ContextBP of
      0,1,2 : begin
              //Quelle est la valeur affichée ?
              RI:=DonneValeurAffiche(codeSession);
              case RI of
               1 : ChpSql:=' QBR_REF1,QBR_REFPRCT1,QBR_OP1,QBR_OPPRCT1 ';
               2 : ChpSql:=' QBR_QTEREF,QBR_QTEREFPRCT,QBR_QTEC,QBR_QTECPRCT ';
               3 : ChpSql:=' QBR_REF2,QBR_REFPRCT2,QBR_OP2,QBR_OPPRCT2 ';
               4 : ChpSql:=' QBR_REF3,QBR_REFPRCT3,QBR_OP3,QBR_OPPRCT3 ';
               5 : ChpSql:=' QBR_REF4,QBR_REFPRCT4,QBR_OP4,QBR_OPPRCT4 ';
               6 : ChpSql:=' QBR_REF5,QBR_REFPRCT5,QBR_OP5,QBR_OPPRCT5 ';
               7 : ChpSql:=' QBR_REF6,QBR_REFPRCT6,QBR_OP6,QBR_OPPRCT6 ';
              end;
            end;
      3 :  begin
             for i:=1 to nbTabAffiche do
               begin               //Tab1 = Val2, Tab2 = Val1, Tab3 = Val3...
                 if i = 1 then ChpSql:=' QBR_QTEREF,QBR_QTEREFPRCT,QBR_QTEC,QBR_QTECPRCT '
                 else if i = 2 then ChpSql:=ChpSql + ',QBR_REF1,QBR_REFPRCT1,QBR_OP1,QBR_OPPRCT1 '
                 else ChpSql:= ChpSql + ',QBR_REF'+IntToStr(i-1)+',QBR_REFPRCT'+IntToStr(i-1)+',QBR_OP'+IntToStr(i-1)+',QBR_OPPRCT'+IntToStr(i-1)+' ';
              end;
           end;
      end;

     Case ContextBp of
      0,1,2 : nbValeurTabD:=4;
      3 : nbValeurTabD:=4*nbTabAffiche;
     end;

     SetLength(TabD,nbValeurTabD);
     SetLength(TabLibCol,nbValeurTabD);

     Case ContextBp of
      0,1,2 : begin
              TabLibCol[0]:=TraduireMemoire('Valeur de référence');
              TabLibCol[1]:=TraduireMemoire('% de référence');
              TabLibCol[2]:=TraduireMemoire('Objectif');
              TabLibCol[3]:=TraduireMemoire('Nouveau % de répartition');

              TabD[0]:='VALEUR_DE_REFERENCE';
              TabD[1]:='PRCT_DE_REFERENCE';
              TabD[2]:='OBJECTIF';
              TabD[3]:='NOUVEAU_PRCT_DE_REPARTITION';
            end;
      3 : begin
            j:=0;

            for i := 0 to nbTabAffiche-1 do
            begin

              TabLibCol[i+j]:=TraduireMemoire('Valeur de référence ('+TabLibValAff[i+1]+')');
              TabLibCol[i+j+1]:=TraduireMemoire('% de référence ('+TabLibValAff[i+1]+')');
              TabLibCol[i+j+2]:=TraduireMemoire('Objectif ('+TabLibValAff[i+1]+')');
              TabLibCol[i+j+3]:=TraduireMemoire('Nouveau % de répartition ('+TabLibValAff[i+1]+')');

              TabD[i+j]:='VALEUR_DE_REFERENCE_'+IntToStr(i+1);
              TabD[i+j+1]:='PRCT_DE_REFERENCE_'+IntToStr(i+1);
              TabD[i+j+2]:='OBJECTIF_'+IntToStr(i+1);
              TabD[i+j+3]:='NOUVEAU_PRCT_DE_REPARTITION_'+IntToStr(i+1);

              j:=j+3;

            end;
          end;
     end;

     ChpColonneD:=TabD[0]+';'+TabD[1]+';'+TabD[2]+';'+TabD[3];

     end
    else
     //-----------------> ORLI
     begin
      if OkSessionObjectif(codeSession)
       then
        begin
         //colonnes à afficher
         RechercheTabAff(TabValAff,TabLibValAff);
         //nombre de colonnes à afficher
         nbTabAffiche:=0;
         for i:=0 to 7 do
          begin
           if TabValAff[i]<>''
            then nbTabAffiche:=nbTabAffiche+1;
          end;

         nbValeurTabD:=4*nbTabAffiche;
         SetLength(TabD,nbValeurTabD);
         SetLength(TabLibCol,nbValeurTabD);
         if SessionDelai(codeSession)
          then
           begin
            nivMax:=nivMax+1;
            TabCodeAxe[NivMax+1]:='DELAI';
           end;

        end
       else
        begin
         if VueCoeff='OUI'
          then
           begin
            nbValeurTabD:=31;
            SetLength(TabD,nbValeurTabD);
            SetLength(TabLibCol,nbValeurTabD);
            okCoeff:=true;
           end;
         if VuePrev='OUI'
          then
           begin
            nivMax:=nivMax+2;
            TabCodeAxe[nivMax]:='ARTICLE';
            TabNumValAxe[nivMax]:='2';
            //niv coloris
            if GestionBPColoris
             then
              begin
               nivMax:=nivMax+1;
               TabCodeAxe[nivMax]:='COLORIS';
               TabNumValAxe[nivMax]:='3';
              end;
            //FS
            if GestionBPFS
             then
              begin
               nivMax:=nivMax+1;
               TabCodeAxe[nivMax]:='FS';
               TabNumValAxe[nivMax]:='4';
              end;
            //niv Magasin
            if GestionBPMagasin
             then
              begin
               nivMax:=nivMax+1;
               TabCodeAxe[nivMax]:='MAG';
               TabNumValAxe[nivMax]:='7';
              end;
            //niv delai
            if SessionDelai(codeSession)
             then
              begin
               nivMax:=nivMax+1;
               TabCodeAxe[nivMax]:='DELAI';
              end;

            nbValeurTabD:=3;
            SetLength(TabD,nbValeurTabD);
            SetLength(TabLibCol,nbValeurTabD);
            OkPrev:=true;
            ChpSql:=' QBR_REALISE,QBR_PREVU,QBR_QTERETENUE ';
            TabD[0]:='QUANTITE_REALISEE';
            TabD[1]:='QUANTITE_PREVUE';
            TabD[2]:='QUANTITE_RETENUE';
            ChpColonneD:=TabD[0]+';'+TabD[1]+';'+TabD[2]+';';

            TabLibCol[0]:=TraduireMemoire('Quantité réalisée');
            TabLibCol[1]:=TraduireMemoire('Quantité prévue');
            TabLibCol[2]:=TraduireMemoire('Quantité retenue');
            nivMax:=ChercheNivMaxSession(codeSession);
           end;
        end;

      for i:=0 to 7 do
       begin
        if TabValAff[i]<>''
         then
          begin
           ChpVirgule:='';
           if ChpSql<>''
            then ChpVirgule:=',';
           nbVal:=copy(TabValAff[i],4,1);
           if nbVal<>'2'
            then
             begin
              if nbVal<>'1'
               then nbVal:=IntToStr(VALEURI(nbVal)-1);
              ChpSql:=ChpSql+ChpVirgule+'QBR_REF'+nbVal+',QBR_REFPRCT'+nbVal+','+
                      'QBR_OP'+nbVal+',QBR_OPPRCT'+nbVal+' ';
             end
            else ChpSql:=ChpSql+ChpVirgule+'QBR_QTEREF,QBR_QTEREFPRCT,QBR_QTEC,QBR_QTECPRCT ';
           codeType:=TraduireMemoire('Prévision');
           if OkSessionObjectif(codeSession)
            then codeType:=TraduireMemoire('Objectif');

           TabD[0+(4*i)]:='VALEUR_DE_REFERENCE'+'_'+UPPERCASE(RemplaceEspaceParTrait(TabLibValAff[i]));
           TabD[1+(4*i)]:='PRCT_DE_REFERENCE'+'_'+UPPERCASE(RemplaceEspaceParTrait(TabLibValAff[i]));
           TabD[2+(4*i)]:=UPPERCASE(codeType)+'_'+UPPERCASE(RemplaceEspaceParTrait(TabLibValAff[i]));
           TabD[3+(4*i)]:='NOUVEAU_PRCT_DE_REPARTITION'+'_'+UPPERCASE(RemplaceEspaceParTrait(TabLibValAff[i]));

           TabLibCol[0+(4*i)]:=TraduireMemoire('Valeur de référence')+' ('+TabLibValAff[i]+')';
           TabLibCol[1+(4*i)]:=TraduireMemoire('% de référence')+' ('+TabLibValAff[i]+')';
           TabLibCol[2+(4*i)]:=codeType+' ('+TabLibValAff[i]+')';
           TabLibCol[3+(4*i)]:=TraduireMemoire('Nouveau % de répartition')+' ('+TabLibValAff[i]+')';

           ChpVirgule:='';
           if ChpColonneD<>''
            then ChpVirgule:=';';
           ChpColonneD:=ChpColonneD+ChpVirgule+TabD[0+(4*i)]+';'+
                        TabD[1+(4*i)]+';'+TabD[2+(4*i)]+';'+TabD[3+(4*i)];
        end;
       end;

      if OkCoeff
       then
        begin
         TabD[0]:='COEFFICIENT';
         TabD[1]:='COEFFICIENT_RETENU';
         TabD[2]:='COEFFICIENT_GENERAL';
         TabD[3]:='NB_CLIENTS_REFERENCE';
         TabD[4]:='NB_CLIENTS_PERDU';
         TabD[5]:='NB_CLIENTS_NOUVEAU';
         TabD[6]:='NB_CLIENTS_VUS';
         TabD[7]:='NB_CLIENTS_RESTE_A_VOIR';
         TabD[8]:='NB_CLIENTS_PROSPECT';
         TabD[9]:='NB_CLIENTS_COURANT';
         TabD[10]:='RESTE_A_VISITER';
         TabD[11]:='XVISIT';
         TabD[12]:='CA_PROSPECT';
         TabD[13]:='QTE_PROSPECT';
         TabD[14]:='NOUVEAU_PROSPECT';
         TabD[15]:='CA_REFERENCE_CLIENTS_VUS';
         TabD[16]:='CA_COURANT_CLIENTS_VUS';
         TabD[17]:='CA_NOUVEAU_CLIENT';
         TabD[18]:='CA_CLIENT_PERDU';
         TabD[19]:='QTE_REFERENCE_CLIENTS_VUS';
         TabD[20]:='QTE_COURANT_CLIENTS_VUS';
         TabD[21]:='QTE_NOUVEAU_CLIENT';
         TabD[22]:='QTE_CLIENT_PERDU';
         TabD[23]:='CA_TOTAL_REFERENCE';
         TabD[24]:='CA_TOTAL_COURANT';
         TabD[25]:='CA_PREVU';
         TabD[26]:='CA_RETENU';
         TabD[27]:='QTE_TOTAL_REFERENCE';
         TabD[28]:='QTE_TOTAL_COURANT';
         TabD[29]:='QTE_PREVU';
         TabD[30]:='QTE_RETENU';

         TabLibCol[0]:=TraduireMemoire('Coefficient');
         TabLibCol[1]:=TraduireMemoire('Coefficient retenu');
         TabLibCol[2]:=TraduireMemoire('Coefficient général');
         TabLibCol[3]:=TraduireMemoire('Nb clients référence');
         TabLibCol[4]:=TraduireMemoire('Nb clients perdu');
         TabLibCol[5]:=TraduireMemoire('Nb clients nouveau');
         TabLibCol[6]:=TraduireMemoire('Nb clients vus');
         TabLibCol[7]:=TraduireMemoire('Nb clients reste à voir');
         TabLibCol[8]:=TraduireMemoire('Nb clients prospect');
         TabLibCol[9]:=TraduireMemoire('Nb clients courant');
         TabLibCol[10]:=TraduireMemoire('Reste à visiter');
         TabLibCol[11]:=TraduireMemoire('XVisit');
         TabLibCol[12]:=TraduireMemoire('CA Prospect');
         TabLibCol[13]:=TraduireMemoire('Qte Prospect');
         TabLibCol[14]:=TraduireMemoire('Nouveau Prospect');
         TabLibCol[15]:=TraduireMemoire('CA référence clients vus');
         TabLibCol[16]:=TraduireMemoire('CA courant clients vus');
         TabLibCol[17]:=TraduireMemoire('CA nouveaux clients');
         TabLibCol[18]:=TraduireMemoire('CA clients perdus');
         TabLibCol[19]:=TraduireMemoire('Qte référence clients vus');
         TabLibCol[20]:=TraduireMemoire('Qte courant clients vus');
         TabLibCol[21]:=TraduireMemoire('Qte nouveaux clients');
         TabLibCol[22]:=TraduireMemoire('Qte clients perdus');
         TabLibCol[23]:=TraduireMemoire('CA total référence');
         TabLibCol[24]:=TraduireMemoire('CA total courant');
         TabLibCol[25]:=TraduireMemoire('CA prévu');
         TabLibCol[26]:=TraduireMemoire('CA retenu');
         TabLibCol[27]:=TraduireMemoire('Qte total référence');
         TabLibCol[28]:=TraduireMemoire('Qte total courant');
         TabLibCol[29]:=TraduireMemoire('Qte prévu');
         TabLibCol[30]:=TraduireMemoire('Qte retenu');

         ChpColonneD:=TabD[0]+';'+TabD[1]+';'+TabD[2]+';'+TabD[3]+';'+TabD[4]+';'+
                      TabD[5]+';'+TabD[6]+';'+TabD[7]+';'+TabD[8]+';'+TabD[9]+';'+
                      TabD[10]+';'+TabD[11]+';'+TabD[12]+';'+TabD[13]+';'+TabD[14]+';'+
                      TabD[15]+';'+TabD[16]+';'+TabD[17]+';'+TabD[18]+';'+TabD[19]+';'+
                      TabD[20]+';'+TabD[21]+';'+TabD[22]+';'+TabD[23]+';'+TabD[24]+';'+
                      TabD[25]+';'+TabD[26]+';'+TabD[27]+';'+TabD[28]+';'+TabD[29]+';'+
                      TabD[30];
        //ajoute chp concernant le coeff
        ChpSql:=ChpSql+'QBR_COEFFCALCUL,QBR_COEFFRETENU,QBR_COEFFGENE,'+
                       'QBR_NBCLTREF,QBR_NBCLTPERDU,QBR_NBCLTNOUVEAU,'+
                       'QBR_NBCLTVU,QBR_NBCLTRESTAVOIR,'+
                       'QBR_NBCLTPROSPECT,QBR_NBCLTC,'+
                       'QBR_RESTEAVISTER,QBR_XVISIT,'+
                       'QBR_PROSPECT,QBR_PROSPECTQTE,QBR_NEWPROSPECT,'+
                       'QBR_VUREF,QBR_VUCOURANT,QBR_NOUVEAU,QBR_PERDU,'+
                       'QBR_VUREFQTE,QBR_VUCOURANTQTE,QBR_NOUVEAUQTE,QBR_PERDUQTE,'+
                       //'QBR_TOTALREF,QBR_TOTALCOURANT,',+
                       'QBR_HISTOCA,QBR_REALISECA,QBR_PREVUCA,QBR_CARETENU,'+
                       'QBR_HISTO,QBR_REALISE,QBR_PREVU,QBR_QTERETENUE';

        end;
     end;
     //ORLI <-----------------

  NbChp:=(2*NivMax);
  SetLength(TabChp,NbChp+1);
  for i:=1 to NivMax do
   begin
    codeAxe:='AXE_'+(RemplaceEspaceParTrait(TabCodeAxe[i+1]));
    TabChp[i-1]:=codeAxe;
    ChpS:=ChpS+codeAxe+';';
   end;
  for i:=NivMax to NivMax*2 do
   begin
    codeAxe:=RemplaceEspaceParTrait(TabCodeAxe[i-NivMax+2]);
    TabChp[i]:='LIBELLE_DE_'+codeAxe;
    ChpS:=ChpS+TabChp[i]+';';
   end;

  TabChp[NbChp]:='NOEUD';
  ChpS:=ChpS+'NOEUD;';

  codeFiltre:=DonneFiltre;

  killTob;
  QTBP:=TQTob.create(TabChp,TabD);


  TFStat(Ecran).LaTOB := QTBP.latob; //TOB
  deblocageMenu(TFStat(ecran));
end ;

procedure TOF_QUFSBPARBREPREV.chargeTob ;
var codefiltre: hString;
begin
  codeFiltre:= DonneFiltre;
  QTBP.latob.clearDetail;
  //remplit la tob
    MetArbreEnLigneNiveau(nivMax,'INITIALISATION',codeSession,ChpSql,codeFiltre,
                        TotalCA1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6,TotalQte,
                        TabCodeAxe,QTBP,okCoeff,okPrev,nbTabAffiche);

  deblocageMenu(TFStat(ecran));
end ;


procedure TOF_QUFSBPARBREPREV.OnLoad ;
begin
  Inherited ;
  chargeTob;
end ;

procedure TOF_QUFSBPARBREPREV.OnArgument (S : String ) ;
var RI:integer;
    titre:hString;
begin
  Inherited ;
  QTBP:= nil;
  TToolBarButton97(GetControl('BTNMODIF')).OnClick := ClickModif;
  TToolBarButton97(GetControl('BTNMODIF')).visible:=false;
  //TTobViewer(getcontrol('TV')).OnLineEnter := LineEnter;
  TTobViewer(getcontrol('TV')).OnDblClick := DblClick;
  codeSession:=TrouveArgument(S,'SESSION','');
  VuePgiDelai:=TrouveArgument(S,'VUEPGIDELAI','');
  VueCoeff:=TrouveArgument(S,'VUECOEFF','');
  VuePrev:=TrouveArgument(S,'VUEPREV','');
  TFStat(Ecran).OnAfterFormShow := LafterShow;
  titre:='Session ';
  TFStat(Ecran).Caption:=titre+': '+codeSession;
  if ContextBP = 1
   then
    begin
     titre:=TraduireMemoire('Contrôle session globale ');
     RI:=DonneValeurAffiche(codeSession);
     case RI of
      1 : TFStat(Ecran).Caption:=titre+TraduireMemoire(' (Prix TTC) : ')+codeSession;
      2 : TFStat(Ecran).Caption:=titre+TraduireMemoire(' (Quantité) : ')+codeSession;
      3 : TFStat(Ecran).Caption:=titre+TraduireMemoire(' (Prix HT) : ')+codeSession;
      4 : TFStat(Ecran).Caption:=titre+TraduireMemoire(' (Prix unitaire HT) : ')+codeSession;
      5 : TFStat(Ecran).Caption:=titre+TraduireMemoire(' (Prix unitaire TTC) : ')+codeSession;
      6 : TFStat(Ecran).Caption:=titre+TraduireMemoire(' (Prix achat HT) : ')+codeSession;
      7 : TFStat(Ecran).Caption:=titre+TraduireMemoire(' (Marge) : ')+codeSession;
     end;
    end;
 if (VueCoeff='OUI') and (not okSessionObjectif(codeSession))
  then
   begin
    TTobViewer(getcontrol('TV')).PopupData := THPopupMenu(getcontrol('POPDET'));
    THPopupMenu(getcontrol('POPDET')).Items[0].OnClick:=ClickDetailClient;
   end;
  creertob;
end ;

procedure TOF_QUFSBPARBREPREV.OnClose ;
begin
  Inherited ;
  killTob;
end ;

procedure TOF_QUFSBPARBREPREV.OnDisplay () ;
begin
  inherited;

  chgColonne;
end;

procedure TOF_QUFSBPARBREPREV.ChgColonne() ;
var axe:integer;
    lib,i,val:integer;
    Libelle : hString;
    Pres: hString;
begin
  Pres:= TFStat(Ecran).FPresentations.Value;
  if Pres<>'' then exit; // on reprend les tittre de la prèsentation
  Libelle := TraduireMemoire('Libellé de ');
  //titre des colonnes axes
  with TTobViewer(getcontrol('TV')) do
   begin
    for i:=0 to nivMax-1 do
     begin
      axe := ColIndex(TabChp[i]);
      if axe >= 0
       then ColCaption[axe]:=DonneLibelleCAxe(structure,TabNumValAxe[i+2],TabCodeAxe[i+2]);
     end;
    for i:=nivMax to nivMax*2-1 do
     begin
      lib := ColIndex(TabChp[i]);
      if lib >= 0
       then ColCaption[lib]:=Libelle+DonneLibelleCAxe(structure,TabNumValAxe[i-nivMax+2],TabCodeAxe[i-nivMax+2]);
     end;
    for i:=0 to nbValeurTabD-1 do
     begin
      Val := ColIndex(TabD[i]);
      if Val >= 0
       then ColCaption[Val]:=TabLibCol[i];
     end;
    if ColIndex('AXE_DELAI') >= 0
     then ColCaption[ColIndex('AXE_DELAI')]:='Délai';
    if ColIndex('LIBELLE_DE_DELAI') >= 0
     then ColCaption[ColIndex('LIBELLE_DE_DELAI')]:='Libellé de délai';
   end;
end ;

procedure TOF_QUFSBPARBREPREV.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_QUFSBPARBREPREV.killTob;
begin
 if QTBP <> nil
  then
   begin
    QTBP.Free;
    QTBP:=nil;
   end;
end;

procedure TOF_QUFSBPARBREPREV.CodeSqlMiseAJourDonnee(RI:integer;retour:hString;
          var codeSql,codeM:hString);
var Chp1,Chp2,Chp3,Chp4,Chp5:hString;
    saival,evolprct,evolval:hString;
begin
 case RI of
  1 : begin
       Chp1:=' QBR_OP1';
       Chp2:=' QBR_SAISI1';
       Chp3:=' QBR_REF1';
       Chp4:=' QBR_EVOLPRCT1';
       Chp5:=' QBR_EVOLVAL1';
       saival:=TrouveArgument(retour,'SAISI1','');
       evolprct:=TrouveArgument(retour,'EVOLPRCT1','');
       evolval:=TrouveArgument(retour,'EVOLVAL1','');
      end;
  2 : begin
       Chp1:=' QBR_QTEC';
       Chp2:=' QBR_SAISIQTE';
       Chp3:=' QBR_QTEREF';
       Chp4:=' QBR_EVOLQTEPRCT';
       Chp5:=' QBR_EVOLQTE';   
       saival:=TrouveArgument(retour,'SAISIQTE','');
       evolprct:=TrouveArgument(retour,'EVOLPRCTQTE','');
       evolval:=TrouveArgument(retour,'EVOLVALQTE','');
      end;
  3 : begin
       Chp1:=' QBR_OP2';
       Chp2:=' QBR_SAISI2';
       Chp3:=' QBR_REF2';
       Chp4:=' QBR_EVOLPRCT2';
       Chp5:=' QBR_EVOLVAL2';
       saival:=TrouveArgument(retour,'SAISI2','');
       evolprct:=TrouveArgument(retour,'EVOLPRCT2','');
       evolval:=TrouveArgument(retour,'EVOLVAL2','');
      end;
  4 : begin
       Chp1:=' QBR_OP3';
       Chp2:=' QBR_SAISI3';
       Chp3:=' QBR_REF3';
       Chp4:=' QBR_EVOLPRCT3';
       Chp5:=' QBR_EVOLVAL3';      
       saival:=TrouveArgument(retour,'SAISI3','');
       evolprct:=TrouveArgument(retour,'EVOLPRCT3','');
       evolval:=TrouveArgument(retour,'EVOLVAL3','');
      end;
  5 : begin
       Chp1:=' QBR_OP4';
       Chp2:=' QBR_SAISI4';
       Chp3:=' QBR_REF4';
       Chp4:=' QBR_EVOLPRCT4';
       Chp5:=' QBR_EVOLVAL4'; 
       saival:=TrouveArgument(retour,'SAISI4','');
       evolprct:=TrouveArgument(retour,'EVOLPRCT4','');
       evolval:=TrouveArgument(retour,'EVOLVAL4','');
      end;
  6 : begin
       Chp1:=' QBR_OP5';
       Chp2:=' QBR_SAISI5';
       Chp3:=' QBR_REF5';
       Chp4:=' QBR_EVOLPRCT5';
       Chp5:=' QBR_EVOLVAL5';        
       saival:=TrouveArgument(retour,'SAISI5','');
       evolprct:=TrouveArgument(retour,'EVOLPRCT5','');
       evolval:=TrouveArgument(retour,'EVOLVAL5','');
      end;
  7 : begin
       Chp1:=' QBR_OP6';
       Chp2:=' QBR_SAISI6';
       Chp3:=' QBR_REF6';
       Chp4:=' QBR_EVOLPRCT6';
       Chp5:=' QBR_EVOLVAL6';    
       saival:=TrouveArgument(retour,'SAISI6','');
       evolprct:=TrouveArgument(retour,'EVOLPRCT6','');
       evolval:=TrouveArgument(retour,'EVOLVAL6','');
      end;
 end;

 //règle de modification
 //si une valeur saisie est renseignée => on met à jour avec cette valeur
 //si pas renseignée et si evolution renseignée => on met à jour avec aX+b
 if (saival<>'')
  then
   begin
    if codesql<>''
     then codeSql:=codeSql+' , ';
    codeSql:=codeSql+Chp1+'='+STRFPOINT(VALEUR(saival))+', '+
             Chp4+'='+STRFPOINT(0)+', '+
             Chp5+'='+STRFPOINT(0)+' ';
    codeM:=Chp2+'='+STRFPOINT(VALEUR(saival))+' ';
   end
  else
   begin
    if codesql<>''
     then codeSql:=codeSql+' , ';
    codeSql:=codeSql+Chp1+'='+Chp3+'+(('+Chp3+'*'+STRFPOINT(VALEUR(evolprct))+')/100) ';
    codeM:=Chp4+'='+STRFPOINT(VALEUR(evolprct))+' ';
    if (evolval='')
     then evolval:='0';
    codeSql:=codeSql+'+ '+STRFPOINT(VALEUR(evolval))+' ';
    codeM:=codeM+' ,'+Chp5+'='+STRFPOINT(VALEUR(evolval))+' ';
   end;
end;


procedure TOF_QUFSBPARBREPREV.MiseAJourDonneeBP(const retour,noeudcourant:hString);
var code,codeSql,noeudpere,codeM,codeMx,coeff:hString;
    Q:TQuery;
    totalniv,totalQte,totalCA2,totalCA3,totalCA4,totalCA5,totalCA6:double;
    RI,i:integer;
    Chp1,Chp2,Chp3,Chp4,Chp5,SqlQte:hString;
    newVal,newQteVal,newCAVal2,newCAVal3,newCAVal4,newCAVal5,newCAVal6:double;
    AncVal,AncQteVal,AncCAVal2,AncCAVal3,AncCAVal4,AncCAVal5,AncCAVal6:double;
    Qte,CA1,CA2,CA3,CA4,CA5,CA6,TotalPrevu,TotalRet,TotalCaRetenu:double;
    OkCoeff:boolean;   
    rapport,somme,QteRet,SommeQte:double;
    ListeQte,ListeQteRes: array[0..19] of double;
    Q2:TQuery;
begin
 codeSql:='';
 //recupere les valeurs
 coeff:=TrouveArgument(retour,'COEFF','');
 SommeQte:=TrouveArgument(retour,'SOMMEQTE','0');
 SqlQte:=TrouveArgument(retour,'SQLQTE','0');
 //cherche le code de la requête
 code:='0';
 codeSql:='';
 codeM:='';    
 codeMx:='';
 RI:=1;
 
 if not BPOkOrli
  then
   begin
     Case ContextBP of
      0,1,2 : begin
              RI:=DonneValeurAffiche(codeSession);
              CodeSqlMiseAJourDonnee(RI,retour,codeSql,codeM);
            end;
      3 :  begin
             for i:=1 to nbTabAffiche do
               begin               //Tab1 = Val2, Tab2 = Val1, Tab3 = Val3...
                 if i = 1 then
                 begin
                   CodeSqlMiseAJourDonnee(2,retour,codeSql,codeM);
                   codeMx:=codeM;
                 end
                 else if i = 2 then
                 begin
                   CodeSqlMiseAJourDonnee(1,retour,codeSql,codeM);
                   codeMx:=codeMx+','+codeM;
                 end
                 else
                 begin
                   CodeSqlMiseAJourDonnee(i,retour,codeSql,codeM);
                   codeMx:=codeMx+','+codeM;
                 end;
              end;
           end;
      end;
   end
  else
   //-----------------> ORLI
   begin
    //CA1
    CodeSqlMiseAJourDonnee(1,retour,codeSql,codeM);
    codeMx:=codeM;
    //CA2
    CodeSqlMiseAJourDonnee(3,retour,codeSql,codeM);
    codeMx:=codeMx+','+codeM;
    //CA3
    CodeSqlMiseAJourDonnee(4,retour,codeSql,codeM);
    codeMx:=codeMx+','+codeM;
    //CA4
    CodeSqlMiseAJourDonnee(5,retour,codeSql,codeM);
    codeMx:=codeMx+','+codeM;
    //CA5
    CodeSqlMiseAJourDonnee(6,retour,codeSql,codeM);
    codeMx:=codeMx+','+codeM;
    //CA6
    CodeSqlMiseAJourDonnee(7,retour,codeSql,codeM);
    codeMx:=codeMx+','+codeM;
    //Qte
    CodeSqlMiseAJourDonnee(2,retour,codeSql,codeM);
    codeMx:=codeMx+','+codeM;
   end;
   //ORLI <-----------------


 OkCoeff:=false;
 if coeff<>''
  then
   begin
    OkCoeff:=true;
    codeSql:=' QBR_PREVU=(QBR_REALISE*'+STRFPOINT(VALEUR(Coeff))+
             '),QBR_PREVUCA=(QBR_REALISECA*'+STRFPOINT(VALEUR(Coeff))+
             '),QBR_CARETENU=(QBR_REALISECA*'+STRFPOINT(VALEUR(Coeff))+
             '),QBR_QTERETENUE=(QBR_REALISE*'+STRFPOINT(VALEUR(Coeff))+
             ') ';
   end;


 if (codeSql='') and SessionInitPrev(codeSession) and (SommeQte<>0)
  then
   begin
   //mise à jour de l'enreg
   MExecuteSql('UPDATE QBPARBRE SET '+SqlQte+',QBR_QTERETENUE="'+STRFPOINT(SommeQte)+'" '+
               ' WHERE QBR_CODESESSION="'+codeSession+
               '" AND QBR_NUMNOEUD="'+NoeudCourant+'" ',
              'BPEclatement (RemplitArbreNiveauEclatement).');

    codeSql:='QBR_QTERETENUE="'+STRFPOINT(SommeQte)+'" ';
  end;


 if codeSql<>''
  then
   begin
    //cherche le noeud pere
    noeudpere:='';
    Q:=MOpenSql('SELECT QBR_NUMNOEUDPERE FROM QBPARBRE '+
                'WHERE QBR_CODESESSION="'+codeSession+
                '" AND QBR_NUMNOEUD="'+NoeudCourant+'"','QUFSBPARBREPREV_TOF (MiseAJourDonneeBP).',true);
    if not Q.eof
     then noeudpere:=Q.fields[0].asString;
    ferme(Q);
    ArbreTotalNiv(codeSession,noeudpere,Qte,CA1,CA2,CA3,CA4,CA5,CA6,TotalPrevu,TotalRet,TotalCaRetenu);
    //récupere les valeurs courantes
    AncVal:=0;
    AncQteVal:=0;
    AncCAVal2:=0;
    AncCAVal3:=0;
    AncCAVal4:=0;
    AncCAVal5:=0;    
    AncCAVal6:=0;
    //new valeur courante
    Q:=OpenSql('SELECT QBR_OP1,QBR_QTEC,'+
               'QBR_OP2,QBR_OP3,QBR_OP4,QBR_OP5,QBR_OP6 '+
               'FROM QBPARBRE '+
               ' WHERE QBR_CODESESSION="'+codeSession+
               '" AND QBR_NUMNOEUD="'+NoeudCourant+'"',true);
    if not Q.eof
     then
      begin
       AncVal:=Q.fields[0].asFloat;
       AncQteVal:=Q.fields[1].asFloat;
       AncCAVal2:=Q.fields[2].asFloat;
       AncCAVal3:=Q.fields[3].asFloat;
       AncCAVal4:=Q.fields[4].asFloat;
       AncCAVal5:=Q.fields[5].asFloat;
       AncCAVal6:=Q.fields[6].asFloat;
      end;
    ferme(Q);

    //met à jour la table
    MExecuteSql('UPDATE QBPARBRE SET '+codeSql+
               ' WHERE QBR_CODESESSION="'+codeSession+
               '" AND QBR_NUMNOEUD="'+NoeudCourant+'"',
               'TOF_QUFSBPARBREPREV (MiseAJourDonneeBP).');
    if codeMx<>''
     then
      //met à jour la table des valeurs de modifications
      MExecuteSql('UPDATE QBPARBRE SET '+codeMx+
               ' WHERE QBR_CODESESSION="'+codeSession+
               '" AND QBR_NUMNOEUD="'+NoeudCourant+'"',
               'TOF_QUFSBPARBREPREV (MiseAJourDonneeBP).');

    //total dernier niveau
    totalniv:=0;
    totalQte:=0;
    totalCA2:=0;
    totalCA3:=0;
    totalCA4:=0;
    totalCA5:=0;
    totalCA6:=0;
    Q:=OpenSql('SELECT SUM(QBR_OP1),SUM(QBR_QTEC),SUM(QBR_OP2),'+
               'SUM(QBR_OP3),SUM(QBR_OP4),SUM(QBR_OP5),SUM(QBR_OP6),'+
               'SUM(QBR_PREVU),SUM(QBR_QTERETENUE),SUM(QBR_CARETENU) '+
               'FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
               '" AND QBR_NUMNOEUDPERE="'+noeudpere+'"',true);
    if not Q.eof
     then
      begin
       totalniv:=Q.fields[0].asFloat;
       totalQte:=Q.fields[1].asFloat;
       totalCA2:=Q.fields[2].asFloat;
       totalCA3:=Q.fields[3].asFloat;
       totalCA4:=Q.fields[4].asFloat;
       totalCA5:=Q.fields[5].asFloat;
       totalCA6:=Q.fields[6].asFloat;
       TotalPrevu:=Q.fields[7].asFloat;
       TotalRet:=Q.fields[8].asFloat;
      end;
    ferme(Q);

    //met à jour les niveaux précédants
    MiseAjourNiveauPrecedent(true,codeSession,NoeudCourant,noeudpere,nivMax,nivMaxSession,totalniv,totalQte,
                             totalCA2,totalCA3,totalCA4,totalCA5,totalCA6,
                             totalniv-CA1,totalQte-Qte,
                             totalCA2-CA2,totalCA3-CA3,
                             totalCA4-CA4,totalCA5-CA5,totalCA6-CA6,
                             TotalPrevu,TotalRet,TotalCaRetenu,
                             coeff,OkCoeff,OkPrev,OkInitNivTaille);

    if totalniv<>0
     then MExecuteSql('UPDATE QBPARBRE SET QBR_OPPRCT1=((QBR_OP1/'+strFPoint4(totalniv)+')*100) '+
                      ' WHERE QBR_CODESESSION="'+codeSession+
                      '" AND QBR_NUMNOEUDPERE="'+noeudpere+'"',
                      'TOF_QUFSBPARBREPREV (MiseAJourDonneeBP).');
    if TotalQte<>0
     then MExecuteSql('UPDATE QBPARBRE SET QBR_QTECPRCT=((QBR_QTEC/'+strFPoint4(TotalQte)+')*100) '+
                      ' WHERE QBR_CODESESSION="'+codeSession+
                      '" AND QBR_NUMNOEUDPERE="'+noeudpere+'"',
                      'TOF_QUFSBPARBREPREV (MiseAJourDonneeBP).');
     if TotalCa2<>0
      then MExecuteSql('UPDATE QBPARBRE SET QBR_OPPRCT2=((QBR_OP2/'+strFPoint4(TotalCa2)+')*100) '+
                      ' WHERE QBR_CODESESSION="'+codeSession+
                      '" AND QBR_NUMNOEUDPERE="'+noeudpere+'"',
                      'TOF_QUFSBPARBREPREV (MiseAJourDonneeBP).');
     if TotalCa3<>0
      then MExecuteSql('UPDATE QBPARBRE SET QBR_OPPRCT3=((QBR_OP3/'+strFPoint4(TotalCa3)+')*100) '+
                      ' WHERE QBR_CODESESSION="'+codeSession+
                      '" AND QBR_NUMNOEUDPERE="'+noeudpere+'"',
                      'TOF_QUFSBPARBREPREV (MiseAJourDonneeBP).');
     if TotalCa4<>0
      then MExecuteSql('UPDATE QBPARBRE SET QBR_OPPRCT4=((QBR_OP4/'+strFPoint4(TotalCa4)+')*100) '+
                      ' WHERE QBR_CODESESSION="'+codeSession+
                      '" AND QBR_NUMNOEUDPERE="'+noeudpere+'"',
                      'TOF_QUFSBPARBREPREV (MiseAJourDonneeBP).');
     if TotalCa5<>0
      then MExecuteSql('UPDATE QBPARBRE SET QBR_OPPRCT5=((QBR_OP5/'+strFPoint4(TotalCa5)+')*100) '+
                      ' WHERE QBR_CODESESSION="'+codeSession+
                      '" AND QBR_NUMNOEUDPERE="'+noeudpere+'"',
                      'TOF_QUFSBPARBREPREV (MiseAJourDonneeBP).');
     if TotalCa6<>0
      then MExecuteSql('UPDATE QBPARBRE SET QBR_OPPRCT6=((QBR_OP6/'+strFPoint4(TotalCa6)+')*100) '+
                      ' WHERE QBR_CODESESSION="'+codeSession+
                      '" AND QBR_NUMNOEUDPERE="'+noeudpere+'"',
                      'TOF_QUFSBPARBREPREV (MiseAJourDonneeBP).');

     //Mise à jour des % de variation
     MExecuteSql('UPDATE QBPARBRE SET  QBR_PRCTVARIATION1=IIF(QBR_REF1=0,100,(QBR_OP1/QBR_REF1)*100),'+
                 'QBR_PRCTVARIATION2=IIF(QBR_REF2=0,100,(QBR_OP2/QBR_REF2)*100),'+
                 'QBR_PRCTVARIATION3=IIF(QBR_REF3=0,100,(QBR_OP3/QBR_REF3)*100),'+
                 'QBR_PRCTVARIATION4=IIF(QBR_REF4=0,100,(QBR_OP4/QBR_REF4)*100),'+
                 'QBR_PRCTVARIATION5=IIF(QBR_REF5=0,100,(QBR_OP5/QBR_REF5)*100),'+
                 'QBR_PRCTVARIATION6=IIF(QBR_REF6=0,100,(QBR_OP6/QBR_REF6)*100),'+
                 'QBR_PRCTVARIATIONQ=IIF(QBR_QTEREF=0,100,(QBR_QTEC/QBR_QTEREF)*100)'+
                 ' WHERE QBR_CODESESSION="'+codeSession+'" ','BPFctArbre (MAJPrctVariation).');  

    if SessionInitPrev(codeSession)
     then
      begin
       //mise à jour de l'enreg
       MExecuteSql('UPDATE QBPARBRE SET QBR_QTERETENUE=QBR_QTEC '+
                  ' WHERE QBR_CODESESSION="'+codeSession+
                  '" AND QBR_NUMNOEUD="'+NoeudCourant+'" ',
                 'TOF_QUFSBPARBREPREV (MiseAJourDonneeBP).');

       rapport:=0;
       somme:=0;
       QteRet:=0;
       Q:=MOpenSql('SELECT QBR_QTET1,QBR_QTET2,QBR_QTET3,QBR_QTET4,QBR_QTET5,'+
                   'QBR_QTET6,QBR_QTET7,QBR_QTET8,QBR_QTET9,QBR_QTET10,'+
                   'QBR_QTET11,QBR_QTET12,QBR_QTET13,QBR_QTET14,QBR_QTET15,'+
                   'QBR_QTET16,QBR_QTET17,QBR_QTET18,QBR_QTET19,QBR_QTET20,QBR_QTERETENUE '+
                   ' FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
                   '" AND QBR_NUMNOEUD="'+NoeudCourant+'"','QUFSBPARBREPREV_TOF (MiseAJourDonneeBP).',true);
       if not Q.eof
        then
         begin
          for i:=0 to 19 do
           begin
            ListeQte[i]:= valeur(Q.fields[i].asString);
            somme:=somme+ListeQte[i];
           end;
          QteRet:=Q.fields[20].asFloat;
         end;
        ferme(Q);
       if somme<>0
        then rapport:=QteRet/somme;

       BPCalculQteListProrata(0,ListeQte,Rapport,ListeQteRes);

       //mise à jour de l'enreg
      MExecuteSql('UPDATE QBPARBRE SET QBR_QTET1='+strfpoint(ListeQteRes[0])+
                  ',QBR_QTET2='+strfpoint(ListeQteRes[1])+',QBR_QTET3='+strfpoint(ListeQteRes[2])+
                  ',QBR_QTET4='+strfpoint(ListeQteRes[3])+',QBR_QTET5='+strfpoint(ListeQteRes[4])+
                  ',QBR_QTET6='+strfpoint(ListeQteRes[5])+',QBR_QTET7='+strfpoint(ListeQteRes[6])+
                  ',QBR_QTET8='+strfpoint(ListeQteRes[7])+',QBR_QTET9='+strfpoint(ListeQteRes[8])+
                  ',QBR_QTET10='+strfpoint(ListeQteRes[9])+',QBR_QTET11='+strfpoint(ListeQteRes[10])+
                  ',QBR_QTET12='+strfpoint(ListeQteRes[11])+',QBR_QTET13='+strfpoint(ListeQteRes[12])+
                  ',QBR_QTET14='+strfpoint(ListeQteRes[13])+',QBR_QTET15='+strfpoint(ListeQteRes[14])+
                  ',QBR_QTET16='+strfpoint(ListeQteRes[15])+',QBR_QTET17='+strfpoint(ListeQteRes[16])+
                  ',QBR_QTET18='+strfpoint(ListeQteRes[17])+',QBR_QTET19='+strfpoint(ListeQteRes[18])+
                  ',QBR_QTET20='+strfpoint(ListeQteRes[19])+
                  ' WHERE QBR_CODESESSION="'+codeSession+
                  '" AND QBR_NUMNOEUD="'+NoeudCourant+'" ',
                 'BPEclatement (RemplitArbreNiveauEclatement).');
      end;
    if SessionDelai(codeSession)
     then
      begin
       //maj du niveau delai
       newVal:=0;
       newQteVal:=0;
       newCAVal2:=0;
       newCAVal3:=0;
       newCAVal4:=0;
       newCAVal5:=0;
       newCAVal6:=0;
       //new valeur courante
       Q:=OpenSql('SELECT QBR_OP1,QBR_QTEC,'+
                  'QBR_OP2,QBR_OP3,QBR_OP4,QBR_OP5,QBR_OP6 '+
                  'FROM QBPARBRE '+
                  ' WHERE QBR_CODESESSION="'+codeSession+
                  '" AND QBR_NUMNOEUD="'+NoeudCourant+'"',true);
       if not Q.eof
        then
         begin
          newVal:=Q.fields[0].asFloat;
          newQteVal:=Q.fields[1].asFloat;
          newCAVal2:=Q.fields[2].asFloat;
          newCAVal3:=Q.fields[3].asFloat;
          newCAVal4:=Q.fields[4].asFloat;
          newCAVal5:=Q.fields[5].asFloat;
          newCAVal6:=Q.fields[6].asFloat;
         end;
       ferme(Q);
       MiseAjourNiveauSuivant(true,NoeudCourant,codeSession,nivMax,nivMaxSession,
                         RI,newVal,newQteVal,
                         newCAVal2,newCAVal3,newCAVal4,newCAVal5,newCAVal6,
                         AncQteVal-newQteVal,AncVal-newVal,AncCAVal2-newCAVal2,
                         AncCAVal3-newCAVal3,AncCAVal4-newCAVal4,
                         AncCAVal5-newCAVal5,AncCAVal6-newCAVal6,0,0,
                         coeff,OkCoeff,okPrev,OkInitNivTaille);
      end;
    recalculTV;
   end;
end;

procedure TOF_QUFSBPARBREPREV.LineEnter(Sender: TObject);
begin
 ligne := TTobViewer(getcontrol('TV')).CurrentRow;
 noeud := TTobViewer(getcontrol('TV')).ColIndex('NOEUD');
 coeffC := TTobViewer(getcontrol('TV')).ColIndex('COEFFICIENT');
 coeffRet := TTobViewer(getcontrol('TV')).ColIndex('COEFFICIENT_RETENU');
end;

procedure TOF_QUFSBPARBREPREV.ClickModif(Sender: TObject);
var retour,NoeudCourant,coeffCS,coeffRetS:hString;
begin
  if ligne >= 0 then
    NoeudCourant:=TTobViewer(getcontrol('TV')).AsString[noeud, ligne];
  if NoeudCourant='' then
  begin
    PGIInfo('Il faut selectionner la ligne à modifier');
    exit;
  end;
  if  SessionValide(codeSession) then
  begin
    PGIInfo('Modification impossible sur une session validée');
    exit;
  end;
  if (not OkSessionObjectif(codeSession)) and okCoeff then
  begin
    coeffRet := TTobViewer(getcontrol('TV')).ColIndex('COEFFICIENT_RETENU');
    coeffCS:=FloatToStr(TTobViewer(getcontrol('TV')).AsDouble[coeffC, ligne]);
    coeffRetS:=FloatToStr(TTobViewer(getcontrol('TV')).AsDouble[coeffRet, ligne]);
    retour:=aglLanceFiche('Q', 'QUFVBPCOEFFMODIF', '', '','COEFF='+coeffCS+';COEFFRET='+coeffRetS);
  end
     else retour:=aglLanceFiche('Q', 'QUFVBPARBREMODIF', '', '',
                          'SESSION='+codeSession+';NOEUD='+NoeudCourant+
                        ';OKCOEFF='+BoolToStr_(okCoeff)+';OKPREV='+BoolToStr_(okPrev));
  if retour<>'' then
  begin
    //maj de qbparbre
    if TrouveArgument(retour,'COEFF','')<>''
      then MExecuteSql('UPDATE QBPARBRE SET QBR_COEFFRETENU="'+STRFPOINT(VALEUR(TrouveArgument(retour,'COEFF','')))+
                  '" WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD="'+NoeudCourant+'"',
                  'TOF_QUFSBPARBREPREV (ClickModif).');
  end
      else exit;
  if retour<>'' then
     MiseAJourDonneeBP(retour,NoeudCourant);
  TFStat(ecran).chercheClick;
end;

procedure TOF_QUFSBPARBREPREV.recalculTV;
begin
  TFStat(ecran).chercheClick;
end;

procedure TOF_QUFSBPARBREPREV.DblClick(Sender: TObject);
begin
 ligne := TTobViewer(getcontrol('TV')).CurrentRow;
 noeud := TTobViewer(getcontrol('TV')).ColIndex('NOEUD');
 coeffC := TTobViewer(getcontrol('TV')).ColIndex('Coefficient');
 ClickModif(sender);
end;

procedure TOF_QUFSBPARBREPREV.Laftershow;
begin
  recalculTV;
end;


{ EVI / Récupération des critères - multiselection}
function TOF_QUFSBPARBREPREV.DonneFiltre:hString;
var codeFiltre,AxeTemp:hString;
    i,Dot,DotTemp:integer;
begin
  Dot :=0;
  codeFiltre:='';

  { EVI / Récupération des critères dans le cas multiselection }
  for i:=1 to 10 do
  begin
    if (GetControlText('EDTAXE'+intToStr(i))<>'') and (Pos(';',GetControlText('EDTAXE'+intToStr(i)))<>0)
    then
    begin
      Dot := VerifPos(true,GetControlText('EDTAXE'+intToStr(i)), ';',Dot);
      AxeTemp := Copy(GetControlText('EDTAXE'+intToStr(i)),1,Dot-2);
      codefiltre:= codefiltre+' AND QBR_VALAXENIV'+intToStr(i)+' in ("'+AxeTemp+'"';
      while Dot <> -1 do
      begin
        if Dot <> -1 then
        begin
          DotTemp := VerifPos( true, GetControlText('EDTAXE'+intToStr(i)), ';', Dot );
          if DotTemp <> -1 then AxeTemp := Copy( GetControlText('EDTAXE'+intToStr(i)), Dot,DotTemp-Dot-1)
          else AxeTemp := Copy( GetControlText('EDTAXE'+intToStr(i)), Dot,length(GetControlText('EDTAXE'+intToStr(i))));
          if LastDelimiter(';',AxeTemp) = length(AxeTemp) then AxeTemp := copy(AxeTemp,1,length(AxeTemp)-1);
          codefiltre:= codefiltre+',"'+AxeTemp+'"';
          Dot := DotTemp;
        end
      end;
      codefiltre:= codefiltre+')';
    end
    else if GetControlText('EDTAXE'+intToStr(i))<>'' then
    if Uppercase(GetControlText('EDTAXE'+intToStr(i)))<>Uppercase(TraduireMemoire('<<Tous>>')) then codefiltre:=codefiltre+' AND QBR_VALAXENIV'+intToStr(i)+' like "'+GetControlText('EDTAXE'+intToStr(i))+'%" ';
  end;
  result:=codeFiltre;
end;

procedure TOF_QUFSBPARBREPREV.ClickDetailClient(Sender: TObject);
var NivMaxL,ligne,i,indexCol,CoeffInd:integer;
    CodeCol,codeTitre,CoeffRetenu:hString;
    TabCodeAxe:array [1..11] of hString;
    TabNumValAxe:array [1..11] of hString;
begin
 //remplit la table de calcul
// RemplitTableQBPCalculTmp(codeSession);


 NivMaxL:=ChercheNivMax(codeSession);

 ChercheTabCodeAxeTabNumValAxe(codeSession,TabCodeAxe,TabNumValAxe);

 //recupere les donnes de la ligne selectionnee
 ligne := TTobViewer(getcontrol('TV')).CurrentRow;

 CoeffInd:=TTobViewer(getcontrol('TV')).ColIndex('COEFFICIENT_RETENU');
 CoeffRetenu:=FloatToStr(TTobViewer(getcontrol('TV')).AsDouble[CoeffInd,ligne]);

 with TTobViewer(getcontrol('TV')) do
  begin
   codeCol:=' QBU_CODESESSION="'+codeSession+'" ';
   for i:=1 to NivMaxL do
    begin
     indexCol:=i-1;
     if indexCol>=0
      then CodeCol:=codeCol+' AND QBU_VALAXECU'+IntToStr(i)+'="'+
                    TTobViewer(getcontrol('TV')).AsString[indexCol,ligne]+'" ';
     codeTitre:=codeTitre+TabCodeAxe[i+1]+' :'+TTobViewer(getcontrol('TV')).AsString[indexCol,ligne]+'  ';
    end;
  end;

 AglLanceFiche('Q','QUFSBPVUECLTDET','','','SESSION='+codeSession+
               ';WHERE='+codeCol+';TITRE='+codeTitre+';COEFF='+CoeffRetenu);
end;

Initialization
  registerclasses ( [ TOF_QUFSBPARBREPREV ] ) ;
end.
