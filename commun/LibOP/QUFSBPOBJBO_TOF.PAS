{***********UNITE*************************************************
Auteur  ...... :                                                                    
Créé le ...... : 13/06/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFSBPOBJBO ()
Mots clefs ... : TOF;QUFSBPOBJBO
*****************************************************************}
Unit QUFSBPOBJBO_TOF ;

Interface

Uses StdCtrls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
     FE_Main,
{$else}
     uTob,
     MainEagl,
{$ENDIF}
     sysutils,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,ugraph,stat,
     LesDatamems,UDatamem,SynScriptBP;

const Mois:array[1..12] of string =
  (
    {1}'jan',
    {2}'fev',
    {3}'mar',
    {4}'avr',
    {5}'mai',
    {6}'jui',
    {7}'jul',
    {8}'aou',
    {9}'sep',
    {10}'oct',
    {11}'nov',
    {12}'dec'
    );

Type
  TOF_QUFSBPOBJBO = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;  
    private      
     QTBP: TQTob;
     edtsession,BPInitialise,codefiltre:string;
     edtDateDeb,edtDateFin:TDateTime;
     TabCodeAxe:array [1..11] of string;
     TabNumValAxe:array [1..11] of string;
     nivMax:integer;
     DmBP: TDatamem;
     procedure DblClick(Sender: TObject);  
     procedure CalculColonnes();    
     procedure killTob;
     procedure DonneFiltre;
  end ;

Implementation


uses CstCommun,UUtil,
     UTobView,BPCubeTmp,QUFGBPGRAPHHPR_TOF,UCtx,BPBasic,BPFctSession;

procedure TOF_QUFSBPOBJBO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFSBPOBJBO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_QUFSBPOBJBO.OnUpdate ;
var enrBP:TEnreg;
    Pseq:TdmPseq;
    MailleS,ChpS:string;
    anM,MoisM,DayM:word;
    Nb,i,NbChp,An,NumSem:integer;
    TabMaille,TabChp:array of string;
    TabChpVal:array of variant;
    TabMailleD:array of double;
    DateDebRef,DateFinRef:TDateTime;
begin
  Inherited ;
  blocageMenu(TFStat(ecran));
  edtsession := THEdit(GetControl('EDTSESSION')).Text;

  if (ControleSession(edtsession)=false)
   then
    begin
     PGIINFO('Le code session doit être renseigné correctement.');
     deblocageMenu(TFStat(ecran));
     exit;
    end;

  //recherche de la période courante et courante-1
  //à partir de la session entrée
  ChercheDateDDateFPeriode(edtsession,edtDateDeb,edtDateFin,DateDebRef,DateFinRef);

  ChercheTabCodeAxeTabNumValAxe(edtSession,TabCodeAxe,TabNumValAxe);

  //cherche le nombre de niveau definit
  nivMax:=ChercheNivMax(edtsession);

  BPInitialise:=SessionBPInitialise(edtSession);

  //initialisation, suivant les coches, pour les colonnes mailles
  Nb:=5;
  SetLength(TabMaille,Nb);
  for i:=1 to Nb do
   begin
    MailleS:=MailleS+';MAILLE'+IntToStr(i);
    TabMaille[i-1]:='MAILLE'+IntToStr(i);
   end;

  //dm
  dm_init(idm_tmpBPB, 0, ['A','A','A','A','A','A','A','A','A','A','A']
                        , true, [rc_BPAXE,rc_BPAXE,rc_BPAXE,rc_BPAXE,rc_BPAXE,
                                 rc_BPAXE,rc_BPAXE,rc_BPAXE,rc_BPAXE,rc_BPAXE,
                                 rc_BPAXE],DmBP);

  //remplit le dm
  CalculColonnes();

  //création de la tob
  //on crée 52 mailles pour avoir 1 année
  killTob;


  { GC/GRC : EV / Contrôle pour session non initialisée }
  if BPInitialise='' then BPInitialise:='0';
  case BPInitialise[1] of
   '2':begin
        NbChp:=(2*NivMax)+1;
        SetLength(TabChp,NbChp+2);
        TabChp[0]:='SEMAINE';
        ChpS:='SEMAINE;';
       end;
   '3':begin   
        NbChp:=(2*NivMax)+1;
        SetLength(TabChp,NbChp+2);
        TabChp[0]:='QUINZAINE';
        ChpS:='QUINZAINE;';
       end;
   '4':begin    
        NbChp:=(2*NivMax)+1;
        SetLength(TabChp,NbChp+2);
        TabChp[0]:='MOIS';
        ChpS:='MOIS;';
       end;
   '5':begin   
        NbChp:=(2*NivMax)+1;
        SetLength(TabChp,NbChp+2);
        TabChp[0]:='MOIS445';
        ChpS:='MOIS445;';
       end;
   '6':begin    
        NbChp:=(2*NivMax)+1;
        SetLength(TabChp,NbChp+2);
        TabChp[0]:='TRIMESTRE';
        ChpS:='TRIMESTRE;';
       end;
   '7':begin 
        NbChp:=(2*NivMax)+1;
        SetLength(TabChp,NbChp+2);
        TabChp[0]:='QUADRIMESTRE';
        ChpS:='QUADRIMESTRE;';
       end;
    else
     begin
        NbChp:=(2*NivMax);
        SetLength(TabChp,NbChp+2);
     end;
  end;
  for i:=1 to NivMax do
   begin
    TabChp[i]:='AXE'+IntToStr(i);
    ChpS:=ChpS+'AXE'+IntToStr(i)+';';
   end;
  for i:=NivMax+1 to NivMax*2 do
   begin
    TabChp[i]:='LIB'+IntToStr(i-NivMax);
    ChpS:=ChpS+'LIB'+IntToStr(i-NivMax)+';';
   end;

  case BPInitialise[1] of
   '2':begin
        TabChp[(NivMax*2)+1]:='NUMERO_SEMAINE';
        ChpS:=ChpS+'NUMERO_SEMAINE;';
       end;
   '3':begin
        TabChp[(NivMax*2)+1]:='LIBELLE_QUINZAINE';
        ChpS:=ChpS+'LIBELLE_QUINZAINE;';
       end;
   '4':begin
        TabChp[(NivMax*2)+1]:='LIBELLE_MOIS';
        ChpS:=ChpS+'LIBELLE_MOIS;';
       end;
   '5':begin
        TabChp[(NivMax*2)+1]:='LIBELLE_MOIS445';
        ChpS:=ChpS+'LIBELLE_MOIS445;';
       end;
   '6':begin
        TabChp[(NivMax*2)+1]:='LIBELLE_TRIMESTRE';
        ChpS:=ChpS+'LIBELLE_TRIMESTRE;';
       end;
   '7':begin
        TabChp[(NivMax*2)+1]:='LIBELLE_QUADRIMESTRE';
        ChpS:=ChpS+'LIBELLE_QUADRIMESTRE;';
       end;
  end;

  Nb:=3;
  SetLength(TabChpVal,NbChp+2);
  QTBP:=TQTob.create(TabChp,TabMaille);
  //remplit tob
  if dm_pseqAll(idm_tmpBPB, Pseq) = 0
   then
    while dm_pseqlect(Pseq, enrBP) = 0 do
     begin
      SetLength(TabMailleD,Nb+2);
      for i:=0 to Nb-1 do
       TabMailleD[i]:=StrToFloat(enrBP.ch(TMPBPB_MAILLE1+i));
      if TabMailleD[1]<>0
       then TabMailleD[3]:=(TabMailleD[2]/TabMailleD[1])*100;
      if TabMailleD[2]<>0
       then TabMailleD[4]:=(TabMailleD[2]/TabMailleD[2])*100;
      TabChpVal[0]:=STRTODATETIME(enrBP.ch(TMPBPB_TYPE));
      for i:=1 to NivMax do
       TabChpVal[i]:=enrBP.ch(TMPBPB_AXE1+i-1);
      for i:=NivMax+1 to NivMax*2 do
       TabChpVal[i]:=enrBP.ch(TMPBPB_LIB1+i-NivMax-1);

      case BPInitialise[1] of
       '2':begin
            NumSem:=NumSemaine(TabChpVal[0],An);
            TabChpVal[(NivMax*2)+1]:=IntToStr(An)+'_'+MetZero(IntToStr(NumSem),2);
           end;
       '3':begin
            TabChpVal[(NivMax*2)+1]:='';
           end;
       '4':begin
            DecodeDate(TabChpVal[0],anM,MoisM,DayM);
            TabChpVal[(NivMax*2)+1]:=MOIS[MoisM]+'_'+IntToStr(anM);
           end;
       '5':begin
            TabChpVal[(NivMax*2)+1]:='';
           end;
       '6':begin
            TabChpVal[(NivMax*2)+1]:='';
           end;
       '7':begin
            TabChpVal[(NivMax*2)+1]:='';
           end;
      end;
      QTBP.addValeur(TabChpVal,TabMailleD);
     end;

  LibereUnDm(idm_tmpBPB);

  TFStat(Ecran).LaTOB := QTBP.latob; //TOB
  TFStat(Ecran).ColNames := ChpS+MailleS; //Colonnes

  deblocageMenu(TFStat(ecran));
end ;

procedure TOF_QUFSBPOBJBO.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_QUFSBPOBJBO.OnArgument (S : String ) ;
begin
  Inherited ;           
  TTobViewer(getcontrol('TV')).OnDblClick := DblClick;
end ;

procedure TOF_QUFSBPOBJBO.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_QUFSBPOBJBO.OnDisplay () ;
var axe1,axe2,axe3,axe4,axe5,axe6,axe7,axe8,axe9,axe10:integer;
    lib1,lib2,lib3,lib4,lib5,lib6,lib7,lib8,lib9,lib10:integer;
    maille1,maille2,maille3,maille4,maille5:integer;
    structure:string;
begin
  inherited;
  structure:='';
  //titre des colonnes axes
  with TTobViewer(getcontrol('TV')) do
   begin
    axe1 := ColIndex('AXE1');
    axe2 := ColIndex('AXE2');
    axe3 := ColIndex('AXE3');
    axe4 := ColIndex('AXE4');
    axe5 := ColIndex('AXE5');
    axe6 := ColIndex('AXE6');
    axe7 := ColIndex('AXE7');
    axe8 := ColIndex('AXE8');
    axe9 := ColIndex('AXE9');
    axe10 := ColIndex('AXE10');

    lib1 := ColIndex('LIB1');
    lib2 := ColIndex('LIB2');
    lib3 := ColIndex('LIB3');
    lib4 := ColIndex('LIB4');
    lib5 := ColIndex('LIB5');
    lib6 := ColIndex('LIB6');
    lib7 := ColIndex('LIB7');
    lib8 := ColIndex('LIB8');
    lib9 := ColIndex('LIB9');
    lib10 := ColIndex('LIB10');

    maille1 := ColIndex('MAILLE1');
    maille2 := ColIndex('MAILLE2');
    maille3 := ColIndex('MAILLE3');  
    maille4 := ColIndex('MAILLE4');
    maille5 := ColIndex('MAILLE5');

    if axe1 > 0
     then ColCaption[axe1]:=DonneLibelleCAxe(structure,TabNumValAxe[1+1],TabCodeAxe[1+1]);
    if axe2 > 0
     then ColCaption[axe2]:=DonneLibelleCAxe(structure,TabNumValAxe[2+1],TabCodeAxe[2+1]);
    if axe3 > 0
     then ColCaption[axe3]:=DonneLibelleCAxe(structure,TabNumValAxe[3+1],TabCodeAxe[3+1]);
    if axe4 > 0
     then ColCaption[axe4]:=DonneLibelleCAxe(structure,TabNumValAxe[4+1],TabCodeAxe[4+1]);
    if axe5 > 0
     then ColCaption[axe5]:=DonneLibelleCAxe(structure,TabNumValAxe[5+1],TabCodeAxe[5+1]);
    if axe6 > 0
     then ColCaption[axe6]:=DonneLibelleCAxe(structure,TabNumValAxe[6+1],TabCodeAxe[6+1]);
    if axe7 > 0
     then ColCaption[axe7]:=DonneLibelleCAxe(structure,TabNumValAxe[7+1],TabCodeAxe[7+1]);
    if axe8 > 0
     then ColCaption[axe8]:=DonneLibelleCAxe(structure,TabNumValAxe[8+1],TabCodeAxe[8+1]);
    if axe9 > 0
     then ColCaption[axe9]:=DonneLibelleCAxe(structure,TabNumValAxe[9+1],TabCodeAxe[9+1]);
    if axe10 > 0
     then ColCaption[axe10]:=DonneLibelleCAxe(structure,TabNumValAxe[10+1],TabCodeAxe[10+1]);

    if lib1 >= 0
     then ColCaption[lib1]:='LIBELLE_DE_'+DonneLibelleCAxe(structure,TabNumValAxe[1+1],TabCodeAxe[1+1]);
    if lib2 >= 0
     then ColCaption[lib2]:='LIBELLE_DE_'+DonneLibelleCAxe(structure,TabNumValAxe[2+1],TabCodeAxe[2+1]);
    if lib3 >= 0
     then ColCaption[lib3]:='LIBELLE_DE_'+DonneLibelleCAxe(structure,TabNumValAxe[3+1],TabCodeAxe[3+1]);
    if lib4 >= 0
     then ColCaption[lib4]:='LIBELLE_DE_'+DonneLibelleCAxe(structure,TabNumValAxe[4+1],TabCodeAxe[4+1]);
    if lib5 >= 0
     then ColCaption[lib5]:='LIBELLE_DE_'+DonneLibelleCAxe(structure,TabNumValAxe[5+1],TabCodeAxe[5+1]);
    if lib6 >= 0
     then ColCaption[lib6]:='LIBELLE_DE_'+DonneLibelleCAxe(structure,TabNumValAxe[6+1],TabCodeAxe[6+1]);
    if lib7 >= 0
     then ColCaption[lib7]:='LIBELLE_DE_'+DonneLibelleCAxe(structure,TabNumValAxe[7+1],TabCodeAxe[7+1]);
    if lib8 >= 0
     then ColCaption[lib8]:='LIBELLE_DE_'+DonneLibelleCAxe(structure,TabNumValAxe[8+1],TabCodeAxe[8+1]);
    if lib9 >= 0
     then ColCaption[lib9]:='LIBELLE_DE_'+DonneLibelleCAxe(structure,TabNumValAxe[9+1],TabCodeAxe[9+1]);
    if lib10 >= 0
     then ColCaption[lib10]:='LIBELLE_DE_'+DonneLibelleCAxe(structure,TabNumValAxe[10+1],TabCodeAxe[10+1]);

    if maille1 >= 0
     then ColCaption[maille1]:='HISTORIQUE';
    if maille2 >= 0
     then ColCaption[maille2]:='OBJECTIFS';
    if maille3 >= 0
     then ColCaption[maille3]:='REALISE';
    if maille4 >= 0
     then ColCaption[maille4]:='%_OBJECTIFS';
    if maille5 >= 0
     then ColCaption[maille5]:='%_HISTORIQUE';

   end;
end ;

procedure TOF_QUFSBPOBJBO.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_QUFSBPOBJBO.DblClick(Sender: TObject);
var ligne:integer;
    LaDate:TDateTime;
    axe1,axe2,axe3,axe4,axe5,axe6,axe7,axe8,axe9,axe10:integer;
    axe1S,axe2S,axe3S,axe4S,axe5S,axe6S,axe7S,axe8S,axe9S,axe10S,codeChp:string;
    Q:TQuery;
    ChpDate:string;
    Histo,Prevu,Realise:double;
begin
 ligne := TTobViewer(getcontrol('TV')).CurrentRow;
 with TTobViewer(getcontrol('TV')) do
  begin
   axe1 := ColIndex('AXE1');
   axe2 := ColIndex('AXE2');
   axe3 := ColIndex('AXE3');
   axe4 := ColIndex('AXE4');
   axe5 := ColIndex('AXE5');
   axe6 := ColIndex('AXE6');
   axe7 := ColIndex('AXE7');
   axe8 := ColIndex('AXE8');
   axe9 := ColIndex('AXE9');
   axe10 := ColIndex('AXE10');
  end;

 if ligne >= 0
  then
   begin
    QTHPR:=TQTob.create(['LADATE'],['CAHISTO','CAPREVU','CAREALISE']);

    codeChp:='';

    axe1S:='';
    if axe1>=0
     then
      begin
       axe1S:=TTobViewer(getcontrol('TV')).AsString[axe1,ligne];
       codeChp:=codeChp+' AND QBQ_VALAXECT1="'+axe1S+'" ';
      end;
    axe2S:='';
    if axe2>=0
     then
      begin
       axe2S:=TTobViewer(getcontrol('TV')).AsString[axe2,ligne];
       codeChp:=codeChp+' AND QBQ_VALAXECT2="'+axe2S+'" ';
      end;
    axe3S:='';
    if axe3>=0
     then
      begin
       axe3S:=TTobViewer(getcontrol('TV')).AsString[axe3,ligne];
       codeChp:=codeChp+' AND QBQ_VALAXECT3="'+axe3S+'" ';
      end;
    axe4S:='';
    if axe4>=0
     then
      begin
       axe4S:=TTobViewer(getcontrol('TV')).AsString[axe4,ligne];
       codeChp:=codeChp+' AND QBQ_VALAXECT4="'+axe4S+'" ';
      end;
    axe5S:='';
    if axe5>=0
     then
      begin
       axe5S:=TTobViewer(getcontrol('TV')).AsString[axe5,ligne];
       codeChp:=codeChp+' AND QBQ_VALAXECT5="'+axe5S+'" ';
      end;
    axe6S:='';
    if axe6>=0
     then
      begin
       axe6S:=TTobViewer(getcontrol('TV')).AsString[axe6,ligne];
       codeChp:=codeChp+' AND QBQ_VALAXECT6="'+axe6S+'" ';
      end;
    axe7S:='';
    if axe7>=0
     then
      begin
       axe7S:=TTobViewer(getcontrol('TV')).AsString[axe7,ligne];
       codeChp:=codeChp+' AND QBQ_VALAXECT7="'+axe7S+'" ';
      end;
    axe8S:='';
    if axe8>=0
     then
      begin
       axe8S:=TTobViewer(getcontrol('TV')).AsString[axe8,ligne];
       codeChp:=codeChp+' AND QBQ_VALAXECT8="'+axe8S+'" ';
      end;
    axe9S:='';
    if axe9>=0
     then
      begin
       axe9S:=TTobViewer(getcontrol('TV')).AsString[axe9,ligne];
       codeChp:=codeChp+' AND QBQ_VALAXECT9="'+axe9S+'" ';
      end;
    axe10S:='';
    if axe10>=0
     then
      begin
       axe10S:=TTobViewer(getcontrol('TV')).AsString[axe10,ligne];
       codeChp:=codeChp+' AND QBQ_VALAXECT10="'+axe10S+'" ';
      end;

   ChpDate:='QBQ_DATECT';
   Q:=MOPenSql('SELECT '+ChpDate+',SUM(QBQ_HISTO),SUM(QBQ_PREVU),SUM(QBQ_REALISE) '+
               'FROM QBPCUBETMP WHERE QBQ_CODESESSION="'+edtsession+
               '" '+codeChp +' GROUP BY '+ChpDate+' ORDER BY '+ChpDate,'',true);
   while not Q.eof do
    begin
     LaDate:=Q.fields[0].asDateTime;
     Histo:=Q.fields[1].asFloat;
     Prevu:=Q.fields[2].asFloat;
     Realise:=Q.fields[3].asFloat;
     QTHPR.addValeur([LaDate],[Histo,Prevu,Realise]);
     Q.next;
    end;
   ferme(Q);

    titreGraph:='Axes : '+axe1S+' '+axe2S+' '+axe3S+' '+axe4S+' '+axe5S+
                ' '+axe6S+' '+axe7S+' '+axe8S+' '+axe9S+' '+axe10S;

    AglLanceFiche('Q', 'QUFGBPGRAPHHPR', '', '', '');
    QTHPR.free;
   end;
end;

procedure TOF_QUFSBPOBJBO.CalculColonnes;
var Q:TQuery;
    dateP:TDatetime;
    j,i:integer;
    TabCodeAxeR,TabLibValAxe:array [1..11] of string;
    enrBP:TEnreg;
    TabMaille:array [1..3] of double;
    codeChp,codeChp2:string;
    CAHisto,CAPrevu,CARealise:double;
    RI:integer;     
    TabCodeAxe,TabNumValAxe:array [1..11] of string;
begin
 DonneFiltre;
 { GC/GRC : EV / Appel fonction modifiée RemplitTableQBPCubeTmpPgi avec
   valeurs nulles en paramètres }
 RemplitTableQBPCubeTmpPgi(edtSession,edtDateDeb,edtDateFin,0,0);

 ChercheTabCodeAxeTabNumValAxe(edtSession,TabCodeAxe,TabNumValAxe);

 for i:=1 to nivMax do
  begin
   if codeChp=''
    then codeChp:=' QBQ_VALAXECT'+IntToStr(i)
    else codeChp:=codeChp+',QBQ_VALAXECT'+IntToStr(i);
  end;
 for i:=1 to nivMax do
  begin
   codeChp:=codeChp+',QBQ_LIBVALAXECT'+intToStr(i);
  end;

 codeChp2:='QBQ_CAHISTO,QBQ_CAPREVU,QBQ_CAREALISE';
 RI:=DonneValeurAffiche(edtSession);
 case RI of
  1 : codeChp2:='QBQ_CAHISTO,QBQ_CAPREVU,QBQ_CAREALISE ';
  2 : codeChp2:='QBQ_HISTO,QBQ_PREVU,QBQ_REALISE ';
  3 : codeChp2:='QBQ_CAHISTO2,QBQ_CAPREVU2,QBQ_CAREALISE2 ';
  4 : codeChp2:='QBQ_CAHISTO3,QBQ_CAPREVU3,QBQ_CAREALISE3 ';
  5 : codeChp2:='QBQ_CAHISTO4,QBQ_CAPREVU4,QBQ_CAREALISE4 ';
  6 : codeChp2:='QBQ_CAHISTO5,QBQ_CAPREVU5,QBQ_CAREALISE5 ';
  7 : codeChp2:='QBQ_CAHISTO6,QBQ_CAPREVU6,QBQ_CAREALISE6 ';
 end;

 Q:=MOpenSql(' SELECT '+codeChp+',QBQ_DATECT,'+codeChp2+
             ' FROM QBPCUBETMP WHERE '+WhereCtx(['qbpcubetmp'])+codefiltre,
             '', true);
 while not Q.eof do
  begin
   //remplit tableau des valeurs
   for j:=0 to nivMax-1 do
    begin
     TabCodeAxeR[j+1]:=Q.fields[j].asString;
     TabLibValAxe[j+1]:=Q.fields[j+nivmax].asString;
    end;

   dateP:=Q.fields[nivMax*2].asDateTime;

   CAHisto:=Q.fields[(nivMax*2)+1].asFloat;
   CAPrevu:=Q.fields[(nivMax*2)+2].asFloat;
   CArealise:=Q.fields[(nivMax*2)+3].asFloat;

   if dm_trouveenr(idm_tmpBPB,[DateTimeToStr(dateP),TabCodeAxeR[1],TabCodeAxeR[2],
                      TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                      TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                      TabCodeAxeR[9],TabCodeAxeR[10]],enrBP)=0
   then
    begin
     dm_modifchcour(idm_tmpBPB,TMPBPB_MAILLE1,
                    FloatToStr(valeur(enrBP.ch(TMPBPB_MAILLE1))+(CAHisto)));
     dm_modifchcour(idm_tmpBPB,TMPBPB_MAILLE2,
                    FloatToStr(valeur(enrBP.ch(TMPBPB_MAILLE2))+(CAPrevu)));
     dm_modifchcour(idm_tmpBPB,TMPBPB_MAILLE3,
                    FloatToStr(valeur(enrBP.ch(TMPBPB_MAILLE3))+(CArealise)));
    end
   else
    begin
     //ajout dans le dm
     //remplit le tableau des mailles
     TabMaille[1]:=CAHisto;
     TabMaille[2]:=CAPrevu;
     TabMaille[3]:=CARealise;
     dm_creaEnreg(idm_tmpBPB,[DateTimeToStr(dateP),TabCodeAxeR[1],TabCodeAxeR[2],
                      TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                      TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                      TabCodeAxeR[9],TabCodeAxeR[10],
                      TabLibValAxe[1],TabLibValAxe[2],
                      TabLibValAxe[3],TabLibValAxe[4],
                      TabLibValAxe[5],TabLibValAxe[6],
                      TabLibValAxe[7],TabLibValAxe[8],
                      TabLibValAxe[9],TabLibValAxe[10],
                      FloatToStr(TabMaille[1]),FloatToStr(TabMaille[2]),
                      FloatToStr(TabMaille[3])]);

     end;
   Q.next
  end;
 ferme(Q);
end;

procedure TOF_QUFSBPOBJBO.killTob;
begin
 if QTBP <> nil
  then
   begin
    QTBP.Free;
    QTBP:=nil;
   end;
end;

{ GC/GRC : EV / Récupération des critères - multiselection }
procedure TOF_QUFSBPOBJBO.DonneFiltre;
var i,Dot,DotTemp:integer;
AxeTemp:string;
edtAxe1,edtAxe2,edtAxe3,edtAxe4,edtAxe5:string;
edtAxe6,edtAxe7,edtAxe8,edtAxe9,edtAxe10:string;
begin
  codeFiltre:='';

  edtAxe1:=THEdit(GetControl('EDTAXE1')).Text;
  edtAxe2:=THEdit(GetControl('EDTAXE2')).Text;
  edtAxe3:=THEdit(GetControl('EDTAXE3')).Text;
  edtAxe4:=THEdit(GetControl('EDTAXE4')).Text;
  edtAxe5:=THEdit(GetControl('EDTAXE5')).Text;
  edtAxe6:=THEdit(GetControl('EDTAXE6')).Text;
  edtAxe7:=THEdit(GetControl('EDTAXE7')).Text;
  edtAxe8:=THEdit(GetControl('EDTAXE8')).Text;
  edtAxe9:=THEdit(GetControl('EDTAXE9')).Text;
  edtAxe10:=THEdit(GetControl('EDTAXE10')).Text;

  Dot :=0;

  //-----------------> ORLI
  if BPOkOrli
  then
  begin
    if edtAxe1<>''
     then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[2]+' like "'+edtAxe1+'%" ';
    if edtAxe2<>''
     then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[3]+' like "'+edtAxe2+'%" ';
    if edtAxe3<>''
     then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[4]+' like "'+edtAxe3+'%" ';
    if edtAxe4<>''
     then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[5]+' like "'+edtAxe4+'%" ';
    if edtAxe5<>''
     then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[6]+' like "'+edtAxe5+'%" ';
    if edtAxe6<>''
     then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[7]+' like "'+edtAxe6+'%" ';
    if edtAxe7<>''
     then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[8]+' like "'+edtAxe7+'%" ';
    if edtAxe8<>''
     then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[9]+' like "'+edtAxe8+'%" ';
    if edtAxe9<>''
     then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[10]+' like "'+edtAxe9+'%" ';
    if edtAxe10<>''
     then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[11]+' like "'+edtAxe10+'%" ';
  end
  //<----------------- ORLI
  else
  begin
    { GC/GRC : EV / Récupération des critères dans le cas multiselection }
    codefiltre:=' AND QBQ_CODESESSION="'+edtSession+'" ';
    for i:=1 to 10 do
    begin
      if (THEdit(GetControl('EDTAXE'+intToStr(i))).Text<>'') and (Pos(';',THEdit(GetControl('EDTAXE'+intToStr(i))).Text)<>0)
      then
      begin
        Dot := VerifPos(true,THEdit(GetControl('EDTAXE'+intToStr(i))).Text, ';',Dot);
        AxeTemp := Copy(THEdit(GetControl('EDTAXE'+intToStr(i))).Text,1,Dot-2);
        codefiltre:= codefiltre+' AND QBQ_VALAXECT'+intToStr(i)+' in ("'+AxeTemp+'"';
        while Dot <> -1 do
        begin
          if Dot <> -1 then
          begin
            DotTemp := VerifPos( true, THEdit(GetControl('EDTAXE'+intToStr(i))).Text, ';', Dot );
            if DotTemp <> -1 then AxeTemp := Copy( THEdit(GetControl('EDTAXE'+intToStr(i))).Text, Dot,DotTemp-Dot-1)
            else AxeTemp := Copy( THEdit(GetControl('EDTAXE'+intToStr(i))).Text, Dot,length(THEdit(GetControl('EDTAXE'+intToStr(i))).Text));
            if LastDelimiter(';',AxeTemp) = length(AxeTemp) then AxeTemp := copy(AxeTemp,1,length(AxeTemp)-1);
            codefiltre:= codefiltre+',"'+AxeTemp+'"';
            Dot := DotTemp;
          end
        end;
      codefiltre:= codefiltre+')';
      end
      else if THEdit(GetControl('EDTAXE'+intToStr(i))).Text<>'' then
      if THEdit(GetControl('EDTAXE'+intToStr(i))).Text<>'<<Tous>>' then codefiltre:=codefiltre+' AND QBQ_VALAXECT'+intToStr(i)+' like "'+THEdit(GetControl('EDTAXE'+intToStr(i))).Text+'%" ';
    end;
  end;
if BPOkOrli then codeFiltre:=codeFiltre+' AND QBQ_CODESESSION="'+edtSession+'" ';
end;

Initialization
  registerclasses ( [ TOF_QUFSBPOBJBO ] ) ; 
end.
