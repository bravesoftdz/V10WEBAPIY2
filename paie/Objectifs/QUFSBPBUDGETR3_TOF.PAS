{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/10/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFSBPBUDGETR3 ()
Mots clefs ... : TOF;QUFSBPBUDGETR3
Modifié le ... : 29/01/2007
Description .. : affichage du presque réalisé pour les sessions liées au budget DA
*****************************************************************}
Unit QUFSBPBUDGETR3_TOF ;

Interface

uses UTOF,HEnt1,Classes,UGraph,UDatamem,Htb97;

const
  TexteMessage : array[1..12] of hString =  (
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
      MoisI:array[1..12] of hString =
  (
    {1}'01',
    {2}'02',
    {3}'03',
    {4}'04',
    {5}'05',
    {6}'06',
    {7}'07',
    {8}'08',
    {9}'09',
    {10}'10',
    {11}'11',
    {12}'12'
    );

    TabValue : array[ 0..6 ] of string = (
    ( 'QBQ_VALAFF2' ),
    ( 'QBQ_VALAFF1' ),
    ( 'QBQ_VALAFF3' ),
    ( 'QBQ_VALAFF4' ),
    ( 'QBQ_VALAFF5' ),
    ( 'QBQ_VALAFF6' ),
    ( 'QBQ_VALAFF7' )
    );


Type
  TOF_QUFSBPBUDGETR3 = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
   private
    QTBP: TQTob;
    structure,edtsession:hString;
    codeFiltre:hString;
    BCherche: TToolbarButton97;
    edtDateDeb,edtDateFin:TDateTime;
    anDeb,NumSemDeb,NivMax:integer;
    anM,moisM,jourM,anDebQ,moisDebQ,QuinzDeb,NumQuinzDeb:word;
    TabCodeAxe:array [1..11] of hString;
    TabNumValAxe:array [1..11] of hString;
    DmBP: TDatamem;
    BPInitialise:hString;
    nbIntervalle:integer;
    InProgress : boolean;
    sTypeDA : string;
    {$IFNDEF MODE}{$IFDEF GCGC} sSessionN, sSessionN1 : string; {$ENDIF}{$ENDIF}  //CEGID-CCMX le 17/11/2006
    procedure killTob;
    function CocheSemaine(): boolean;
    function CocheQuinzaine(): boolean;
    function CocheMois(): boolean;
    function CocheMois445(): boolean;
    function DonneNomColonne(datePivot:TDateTime;NumSemDeb:integer;
             anDeb,moisDeb:word):integer;
    procedure CalculColonnes(OkDelai:boolean);
    procedure DblClick(Sender: TObject);
    procedure DonneFiltre;
    procedure RemplitTV(Base,Chp,Chp2,filtre:string;ValAff:integer;OKDelay:boolean;TabLib:array of hString);
  end ;

Implementation

Uses Windows,uTob,sysutils,HMsgBox,stat,UCtx,HCtrls,
     StrUtils,ed_tools,CstCommun,UTobView,utilpgi,Controls,
     {$IFNDEF EAGLCLIENT}FE_Main,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     {$else}eMul,MainEagl,{$ENDIF}
     {$IFNDEF MODE}{$IFDEF GCGC}paramsoc,{$ENDIF}{$ENDIF}
     LesDatamems,SynScriptBP,UUtil,BPCubeTmp,
     QUFGBPGRAPHHPR_TOF,BPBasic,BPFctSession,BPUtil;

procedure TOF_QUFSBPBUDGETR3.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFSBPBUDGETR3.OnDelete ;
begin
  Inherited ;
end ;

//cherche le numéro de la maille
function TOF_QUFSBPBUDGETR3.DonneNomColonne(datePivot:TDateTime;
         NumSemDeb:integer;anDeb,moisDeb:word):integer;
var anPP,NumSemP:integer;
    anP,moisP,jourP,quinzP,numQuinzP:word;
begin

 result:=12;
 //pour maille semaine
 if CocheSemaine
  then
   begin
    NumSemP:= numsemaine(datePivot,anPP);
    if (NumSemP>=NumSemDeb)
     then result:=NumSemP - NumSemDeb + 1
     else result:=(52 - NumSemDeb + 1)+numSemP;
   end;
 //pour maille mois
 if CocheMois
  then
   begin
    decodeDate(datePivot,anP,moisP,jourP);
    if (anP=anDeb)
     then result:=moisP - moisDeb + 1
     else
      if (anP>=anDeb)
       then result:=(12*(anP-anDeb) - moisDeb + 1 + moisP);
   end;
 //pour maille quinzaine
 if CocheQuinzaine
  then
   begin
    Date_AnNumQuinz(datePivot,anP,moisP,quinzP,numQuinzP);
    if (anP=anDebQ)
     then result:=numQuinzP-NumQuinzDeb+1
     else result:=(24*(anP-anDebQ) - NumQuinzDeb + 1 + numQuinzP);
   end;
end;


//remplit dm
//pour fiche stat
//à partir de table cube tmp
procedure TOF_QUFSBPBUDGETR3.CalculColonnes(OkDelai:boolean);
var DateEdDeb,DateEdFin:TDatetime;
    i:integer;
    codeChp,codeChp2:hString;
    NbValAff,RI,nivMaxSession:integer;
    TabLibelle:array [0..7] of hString;
    LimitTdB,ValueAxe:hString;
    sSocietes,sSocGroupe,sBase:string;
begin
  NbValAff := 0;
  sTypeDA := ''; //CEGID-CCMX le 17/11/2006
  DonneFiltre;
  if BPOkOrli
  then //-----------------> ORLI
    RemplitTableQBPCubeTmpOrli(edtSession,edtDateDeb,edtDateFin)
       //ORLI <-----------------
  { EVI / Appel fonction modifiée RemplitTableQBPCubeTmpPgi avec
    valeurs nulles en paramètres }
  else
  begin
    {$IFNDEF MODE}
    {$IFDEF GCGC}
    //CEGID-CCMX le 17/11/2006 DEBUT
    sSessionN := GetParamsocSecur('SO_GCCODESESSIONN','');
    sSessionN1 := GetParamsocSecur('SO_GCCODESESSIONN1','');
    // le 16/01/2007 : en fait on limite pas seulement au session N ou N+1 mais aux
    // sessions utilisées
    if (sSessionN = edtSession) or (sSessionN1 = edtSession)
    or (ExisteSQL('SELECT DA_CODESESSION FROM PIECEDA WHERE DA_CODESESSION="' + edtSession+'"'))
    then sTypeDA:='TYPEDA';
    //CEGID-CCMX le 17/11/2006 FIN
    {$ENDIF}
    {$ENDIF}

    //Récupère les dates pour la génération du cube
    if (GetControlText('DATEDEB')='') or (GetControlText('DATEFIN')='') then
    begin
      DateEdDeb := 0;
      DateEdFin := 0;
    end else
    begin
      DateEdDeb := StrToDate(GetControlText('DATEDEB'));
      DateEdFin := StrToDate(GetControlText('DATEFIN'));
    end;

    if DateEdDeb = 0 then DateEdDeb := edtDateDeb;
    if DateEdFin = 0 then DateEdFin := edtDateFin;

     nivMaxSession:=ChercheNivMax(EdtSession);
    //Récupère les restrictions pour la génération du cube
    for i:=1 to nivMaxSession do
    begin
       ValueAxe := GetControlText('EDTAXE'+IntToStr(i));
       ValueAxe:= AnsiReplaceText(ValueAxe, ';', '","');
       if LimitTdB = '' then LimitTdB := '("'+ValueAxe+'")'
       else LimitTdB := LimitTdB + '@@' + '("'+ValueAxe+'")'
    end;

    { EVI / Génération du cube avec restriction utilisateur }
    sSocietes := GetBasesMS(THValComboBox(GetControl('MULTIDOSSIER')).Value, true);
    if sSocietes <> '' then
    begin
      sSocGroupe := sSocietes;
      sBase := ReadTokenSt(sSocGroupe) ;
      repeat
        RemplitTableQBPCubeTmpPgi(edtSession,sBase,LimitTdB,edtDateDeb,edtDateFin,DateEdDeb,DateEdFin);
      sBase  := ReadTokenSt(sSocGroupe) ;
      until (sBase = '');
    end else RemplitTableQBPCubeTmpPgi(edtSession,'',LimitTdB,edtDateDeb,edtDateFin,DateEdDeb,DateEdFin);
  end;

  if not BPOkOrli then
  begin
    for i:=1 to nivMax do
    begin
      if codeChp='' then codeChp:=' QBQ_VALAXECT'+IntToStr(i)
      else codeChp:=codeChp+',QBQ_VALAXECT'+IntToStr(i);
    end;
    for i:=1 to nivMax do codeChp:=codeChp+',QBQ_LIBVALAXECT'+intToStr(i);
  end
  else
  begin //-----------------> ORLI
    for i:=1 to nivMax do
    begin
      if codeChp='' then codeChp:=' QBQ_VALAXECT'+TabNumValAxe[i+1]
      else codeChp:=codeChp+',QBQ_VALAXECT'+TabNumValAxe[i+1];
    end;
    for i:=1 to nivMax do codeChp:=codeChp+',QBQ_LIBVALAXECT'+TabNumValAxe[i+1];
   end; //ORLI <-----------------

  codeChp2:='QBQ_CAHISTO,QBQ_CAPREVU,QBQ_CAREALISE';
  if not BPOkOrli then
  begin
    RI:=DonneValeurAffiche(edtSession);
    Case ContextBP of
      0,1,2 : begin //Mode-GC-Compta
              case RI of
                1 : codeChp2:='QBQ_CAHISTO,QBQ_CAPREVU,QBQ_CAREALISE ';
                2 : codeChp2:='QBQ_HISTO,QBQ_PREVU,QBQ_REALISE ';
                //CEGID-CCMX le 17/11/2006 DEBUT
                3 : if (sTypeDA = 'TYPEDA') then codeChp2:='QBQ_CAHISTO2,QBQ_CAPREVU2,QBQ_CAREALISE2,QBQ_CAPREAL2 '
                                            else codeChp2:='QBQ_CAHISTO2,QBQ_CAPREVU2,QBQ_CAREALISE2 ';
               //CEGID-CCMX le 17/11/2006 FIN
                4 : codeChp2:='QBQ_CAHISTO3,QBQ_CAPREVU3,QBQ_CAREALISE3 ';
                5 : codeChp2:='QBQ_CAHISTO4,QBQ_CAPREVU4,QBQ_CAREALISE4 ';
                6 : codeChp2:='QBQ_CAHISTO5,QBQ_CAPREVU5,QBQ_CAREALISE5 ';
                7 : codeChp2:='QBQ_CAHISTO6,QBQ_CAPREVU6,QBQ_CAREALISE6 ';
              end;
            end;
      3 : begin //Paie
            LibValAff(edtSession,TabLibelle);
            NbValAff := StrToInt(TabLibelle[0]);
            codeChp2 := 'QBQ_NATVALAFF';
            for i := 1 to NbValAff do codeChp2:=codeChp2 + ',' + TabValue[i-1];
          end;
    end; //CASE
  end
  else
  begin  //-----------------> ORLI
    codeChp2:='QBQ_CAHISTO,QBQ_CAPREVU,QBQ_CAREALISE,'+
              'QBQ_HISTO,QBQ_PREVU,QBQ_REALISE';
  end;  //ORLI <-----------------





  if sSocietes <> '' then
  begin
    sSocGroupe := sSocietes;
    sBase := ReadTokenSt(sSocGroupe) ;
    repeat
      RemplitTV(sBase,codeChp,codeChp2,codefiltre,NbValAff,OKDelai,TabLibelle);
    sBase  := ReadTokenSt(sSocGroupe) ;
    until (sBase = '');
  end else RemplitTV('',codeChp,codeChp2,codefiltre,NbValAff,OKDelai,TabLibelle);

  FiniMoveProgressForm;
  InProgress := false;
end;


procedure TOF_QUFSBPBUDGETR3.RemplitTV(Base,Chp,Chp2,filtre:string;ValAff:integer;OKDelay:boolean;TabLib:array of hString);
var QteHisto,QtePrevu,QteRealise:double;
    CAHisto,CAPrevu,CARealise,CAPReal:double; //CEGID-CCMX ajout CAPReal
    Q:TQuery;
    diffDate,i,j,k:integer;
    TabCodeAxeR,TabLibValAxe:array [1..11] of hString;
    dateP:TDatetime;
    TabHisto,TabPrevu,TabRealiseN,TabRealiseNm : array [0..6] of double;
    OkCont:boolean;
    enrBP:TEnreg;
    TabMailleHisto,TabMaillePrevu,tabMailleRealise,tabMailleRealiseN,TabMaillePReal:array [1..54] of double; //CEGID-CCMX ajout TabMaillePReal
    TabMailleQteHisto,TabMailleQtePrevu,TabMailleQteRealise:array [1..54] of double;
    Enreg:boolean;
begin
  QteHisto:=0;
  QtePrevu:=0;
  QteRealise:=0;
  { EVI / Ajout de QBQ_UTILISATEUR}
  Q:=MOpenSql(' SELECT '+Chp+',QBQ_DATECT,'+Chp2+
              ' FROM QBPCUBETMP WHERE '+WhereCtx(['qbpcubetmp'])+filtre+
              ' AND QBQ_UTILISATEUR="'+V_PGI.User +'"',
              'QUFSBPBUDGETR3_TOF (CalculColonnes).', true);

  while not Q.eof do
  begin
    //remplit tableau des valeurs
    for j:=0 to nivMax-1 do
    begin
      TabCodeAxeR[j+1]:=Q.fields[j].asString;
      TabLibValAxe[j+1]:=Q.fields[j+nivmax].asString;
    end;

    CAHisto := 0;
    CAPrevu := 0;
    CArealise := 0;

    dateP:=Q.fields[nivMax*2].asDateTime;
    Case ContextBP of
      0,1,2 : begin //Mode-GC-Compta
              CAHisto:=Q.fields[(nivMax*2)+1].asFloat;
              CAPrevu:=Q.fields[(nivMax*2)+2].asFloat;
              CArealise:=Q.fields[(nivMax*2)+3].asFloat;
            end;
      3 : begin //Paie
            for i := 1 to ValAff do
            begin
              TabHisto[i-1] := 0;
              TabPrevu[i-1] := 0;
              TabRealiseN[i-1] := 0;
              TabRealiseNm[i-1] := 0;
            end;
            if UpperCase(Q.fields[(nivMax*2)+1].asString) = UpperCase('Théorique') then
            for i := 1 to ValAff do TabHisto[i-1] := Q.fields[(nivMax*2)+1+i].asFloat;
            if UpperCase(Q.fields[(nivMax*2)+1].asString) = UpperCase('Budget N') then
            for i := 1 to ValAff do TabPrevu[i-1] := Q.fields[(nivMax*2)+1+i].asFloat;
            if UpperCase(Q.fields[(nivMax*2)+1].asString) = UpperCase('Réalisé N') then
            for i := 1 to ValAff do TabRealiseN[i-1] := Q.fields[(nivMax*2)+1+i].asFloat;
            if UpperCase(Q.fields[(nivMax*2)+1].asString) = UpperCase('Réalisé N-1') then
            for i := 1 to ValAff do TabRealiseNm[i-1] := Q.fields[(nivMax*2)+1+i].asFloat;
          end;
    end; //CASE

    CAPReal := 0;
    if (sTypeDA = 'TYPEDA') then CAPReal := Q.fields[(nivMax*2)+4].asFloat;  //CEGID-CCMX le 17/11/2006
    if BPOkOrli then
    begin //-----------------> ORLI
      QteHisto:=Q.fields[(nivMax*2)+4].asFloat;
      QtePrevu:=Q.fields[(nivMax*2)+5].asFloat;
      QteRealise:=Q.fields[(nivMax*2)+6].asFloat;
    end; //ORLI <-----------------


    OkCont:=false;
    diffDate:=22;
    //si la date appartient bien à l'intervalle entré
    if DateAppartIntervalle(dateP,edtDateDeb,edtDateFin) and OKDelay then
    begin
      //cherche le numéro de la colonne maille
      //on ajoute +21 pc les chps mailles dans le Dm commencent à 22
      diffDate:=DonneNomColonne(dateP,NumSemDeb,anM,moisM)+21;

      //ajout dans le dm
      if (diffDate<=74) and (diffDate>21) then OkCont:=true;
    end;
    if OkDelay=false then OkCont:=true;

    if OkCont then
    begin
      Case ContextBp of
      0,1,2 : begin
              if dm_trouveenr(idm_tmpBPB,['1-HISTO',TabCodeAxeR[1],TabCodeAxeR[2],
                              TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                              TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                              TabCodeAxeR[9],TabCodeAxeR[10]],enrBP)=0
              then
              begin
                dm_modifchcour(idm_tmpBPB,diffdate,FloatToStr(valeur(enrBP.ch(diffdate))+(CAHisto)));

                if dm_trouveenr(idm_tmpBPB,['2-PREVU',TabCodeAxeR[1],TabCodeAxeR[2],
                                TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                TabCodeAxeR[9],TabCodeAxeR[10]],enrBP)=0
                then dm_modifchcour(idm_tmpBPB,diffdate,FloatToStr(valeur(enrBP.ch(diffdate))+(CAPrevu)));

                if dm_trouveenr(idm_tmpBPB,['3-REALISE',TabCodeAxeR[1],TabCodeAxeR[2],
                                TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                TabCodeAxeR[9],TabCodeAxeR[10]],enrBP)=0
                then dm_modifchcour(idm_tmpBPB,diffdate,FloatToStr(valeur(enrBP.ch(diffdate))+(CARealise)));

                //CEGID-CCMX le 17/11/2006 DEBUT
                if (sTypeDA = 'TYPEDA') then
                begin
                  if dm_trouveenr(idm_tmpBPB,['4-PRESQUE REALISE',TabCodeAxeR[1],TabCodeAxeR[2],
                                  TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                   TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                  TabCodeAxeR[9],TabCodeAxeR[10]],enrBP)=0
                  then dm_modifchcour(idm_tmpBPB,diffdate,FloatToStr(valeur(enrBP.ch(diffdate))+(CARealise)));
                end;
                //CEGID-CCMX le 17/11/2006 FIN

                if BPOkOrli then
                begin //-----------------> ORLI
                  if dm_trouveenr(idm_tmpBPB,['4-QTEHISTO',TabCodeAxeR[1],TabCodeAxeR[2],
                                  TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                  TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                  TabCodeAxeR[9],TabCodeAxeR[10]],enrBP)=0
                  then dm_modifchcour(idm_tmpBPB,diffdate,
                                      FloatToStr(valeur(enrBP.ch(diffdate))+(QteHisto)));
                  if dm_trouveenr(idm_tmpBPB,['5-QTEPREVU',TabCodeAxeR[1],TabCodeAxeR[2],
                                  TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                  TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                  TabCodeAxeR[9],TabCodeAxeR[10]],enrBP)=0
                  then dm_modifchcour(idm_tmpBPB,diffdate,
                                      FloatToStr(valeur(enrBP.ch(diffdate))+(QtePrevu)));
                  if dm_trouveenr(idm_tmpBPB,['6-QTEREALISE',TabCodeAxeR[1],TabCodeAxeR[2],
                                  TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                  TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                  TabCodeAxeR[9],TabCodeAxeR[10]],enrBP)=0
                  then dm_modifchcour(idm_tmpBPB,diffdate,
                                      FloatToStr(valeur(enrBP.ch(diffdate))+(QteRealise)));
                end; //ORLI <-----------------
              end
              else
              begin
                //ajout dans le dm
                //remplit le tableau des mailles
                for k:=1 to 53 do
                begin
                  TabMailleHisto[k]:=0;
                  TabMaillePrevu[k]:=0;
                  TabMailleRealise[k]:=0;
                  if (sTypeDA = 'TYPEDA') then TabMaillePReal[k]:=0;    //CEGID-CCMX  le 17/11/2006
                  TabMailleQteHisto[k]:=0;
                  TabMailleQtePrevu[k]:=0;
                  TabMailleQteRealise[k]:=0;
                  if k=diffdate-21 then
                  begin
                    TabMailleHisto[k]:=CAHisto;
                    TabMaillePrevu[k]:=CAPrevu;
                    TabMailleRealise[k]:=CARealise;
                    if (sTypeDA = 'TYPEDA') then TabMaillePReal[k]:=CAPReal;  //CEGID-CCMX le 17/11/2006
                    TabMailleQteHisto[k]:=QteHisto;
                    TabMailleQtePrevu[k]:=QtePrevu;
                    TabMailleQteRealise[k]:=QteRealise;
                  end;
                end;
                dm_creaEnreg(idm_tmpBPB,['1-HISTO',TabCodeAxeR[1],TabCodeAxeR[2],
                                         TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                         TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                         TabCodeAxeR[9],TabCodeAxeR[10],
                                         TabLibValAxe[1],TabLibValAxe[2],
                                         TabLibValAxe[3],TabLibValAxe[4],
                                         TabLibValAxe[5],TabLibValAxe[6],
                                         TabLibValAxe[7],TabLibValAxe[8],
                                         TabLibValAxe[9],TabLibValAxe[10],
                                         FloatToStr(TabMailleHisto[1]),FloatToStr(TabMailleHisto[2]),
                                         FloatToStr(TabMailleHisto[3]),FloatToStr(TabMailleHisto[4]),
                                         FloatToStr(TabMailleHisto[5]),FloatToStr(TabMailleHisto[6]),
                                         FloatToStr(TabMailleHisto[7]),FloatToStr(TabMailleHisto[8]),
                                         FloatToStr(TabMailleHisto[9]),FloatToStr(TabMailleHisto[10]),
                                         FloatToStr(TabMailleHisto[11]),FloatToStr(TabMailleHisto[12]),
                                         FloatToStr(TabMailleHisto[13]),FloatToStr(TabMailleHisto[14]),
                                         FloatToStr(TabMailleHisto[15]),FloatToStr(TabMailleHisto[16]),
                                         FloatToStr(TabMailleHisto[17]),FloatToStr(TabMailleHisto[18]),
                                         FloatToStr(TabMailleHisto[19]),FloatToStr(TabMailleHisto[20]),
                                         FloatToStr(TabMailleHisto[21]),FloatToStr(TabMailleHisto[22]),
                                         FloatToStr(TabMailleHisto[23]),FloatToStr(TabMailleHisto[24]),
                                         FloatToStr(TabMailleHisto[25]),FloatToStr(TabMailleHisto[26]),
                                         FloatToStr(TabMailleHisto[27]),FloatToStr(TabMailleHisto[28]),
                                         FloatToStr(TabMailleHisto[29]),FloatToStr(TabMailleHisto[30]),
                                         FloatToStr(TabMailleHisto[31]),FloatToStr(TabMailleHisto[32]),
                                         FloatToStr(TabMailleHisto[33]),FloatToStr(TabMailleHisto[34]),
                                         FloatToStr(TabMailleHisto[35]),FloatToStr(TabMailleHisto[36]),
                                         FloatToStr(TabMailleHisto[37]),FloatToStr(TabMailleHisto[38]),
                                         FloatToStr(TabMailleHisto[39]),FloatToStr(TabMailleHisto[40]),
                                         FloatToStr(TabMailleHisto[41]),FloatToStr(TabMailleHisto[42]),
                                         FloatToStr(TabMailleHisto[43]),FloatToStr(TabMailleHisto[44]),
                                         FloatToStr(TabMailleHisto[45]),FloatToStr(TabMailleHisto[46]),
                                         FloatToStr(TabMailleHisto[47]),FloatToStr(TabMailleHisto[48]),
                                         FloatToStr(TabMailleHisto[49]),FloatToStr(TabMailleHisto[50]),
                                         FloatToStr(TabMailleHisto[51]),FloatToStr(TabMailleHisto[52]),
                                         FloatToStr(TabMailleHisto[53])]);
                dm_creaEnreg(idm_tmpBPB,['2-PREVU',TabCodeAxeR[1],TabCodeAxeR[2],
                                         TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                         TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                         TabCodeAxeR[9],TabCodeAxeR[10],
                                         TabLibValAxe[1],TabLibValAxe[2],
                                         TabLibValAxe[3],TabLibValAxe[4],
                                         TabLibValAxe[5],TabLibValAxe[6],
                                         TabLibValAxe[7],TabLibValAxe[8],
                                         TabLibValAxe[9],TabLibValAxe[10],
                                         FloatToStr(TabMaillePrevu[1]),FloatToStr(TabMaillePrevu[2]),
                                         FloatToStr(TabMaillePrevu[3]),FloatToStr(TabMaillePrevu[4]),
                                         FloatToStr(TabMaillePrevu[5]),FloatToStr(TabMaillePrevu[6]),
                                         FloatToStr(TabMaillePrevu[7]),FloatToStr(TabMaillePrevu[8]),
                                         FloatToStr(TabMaillePrevu[9]),FloatToStr(TabMaillePrevu[10]),
                                         FloatToStr(TabMaillePrevu[11]),FloatToStr(TabMaillePrevu[12]),
                                         FloatToStr(TabMaillePrevu[13]),FloatToStr(TabMaillePrevu[14]),
                                         FloatToStr(TabMaillePrevu[15]),FloatToStr(TabMaillePrevu[16]),
                                         FloatToStr(TabMaillePrevu[17]),FloatToStr(TabMaillePrevu[18]),
                                         FloatToStr(TabMaillePrevu[19]),FloatToStr(TabMaillePrevu[20]),
                                         FloatToStr(TabMaillePrevu[21]),FloatToStr(TabMaillePrevu[22]),
                                         FloatToStr(TabMaillePrevu[23]),FloatToStr(TabMaillePrevu[24]),
                                         FloatToStr(TabMaillePrevu[25]),FloatToStr(TabMaillePrevu[26]),
                                         FloatToStr(TabMaillePrevu[27]),FloatToStr(TabMaillePrevu[28]),
                                         FloatToStr(TabMaillePrevu[29]),FloatToStr(TabMaillePrevu[30]),
                                         FloatToStr(TabMaillePrevu[31]),FloatToStr(TabMaillePrevu[32]),
                                         FloatToStr(TabMaillePrevu[33]),FloatToStr(TabMaillePrevu[34]),
                                         FloatToStr(TabMaillePrevu[35]),FloatToStr(TabMaillePrevu[36]),
                                         FloatToStr(TabMaillePrevu[37]),FloatToStr(TabMaillePrevu[38]),
                                         FloatToStr(TabMaillePrevu[39]),FloatToStr(TabMaillePrevu[40]),
                                         FloatToStr(TabMaillePrevu[41]),FloatToStr(TabMaillePrevu[42]),
                                         FloatToStr(TabMaillePrevu[43]),FloatToStr(TabMaillePrevu[44]),
                                         FloatToStr(TabMaillePrevu[45]),FloatToStr(TabMaillePrevu[46]),
                                         FloatToStr(TabMaillePrevu[47]),FloatToStr(TabMaillePrevu[48]),
                                         FloatToStr(TabMaillePrevu[49]),FloatToStr(TabMaillePrevu[50]),
                                         FloatToStr(TabMaillePrevu[51]),FloatToStr(TabMaillePrevu[52]),
                                         FloatToStr(TabMaillePrevu[53])]);
                dm_creaEnreg(idm_tmpBPB,['3-REALISE',TabCodeAxeR[1],TabCodeAxeR[2],
                                         TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                         TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                         TabCodeAxeR[9],TabCodeAxeR[10],
                                         TabLibValAxe[1],TabLibValAxe[2],
                                         TabLibValAxe[3],TabLibValAxe[4],
                                         TabLibValAxe[5],TabLibValAxe[6],
                                         TabLibValAxe[7],TabLibValAxe[8],
                                         TabLibValAxe[9],TabLibValAxe[10],
                                         FloatToStr(TabMailleRealise[1]),FloatToStr(TabMailleRealise[2]),
                                         FloatToStr(TabMailleRealise[3]),FloatToStr(TabMailleRealise[4]),
                                         FloatToStr(TabMailleRealise[5]),FloatToStr(TabMailleRealise[6]),
                                         FloatToStr(TabMailleRealise[7]),FloatToStr(TabMailleRealise[8]),
                                         FloatToStr(TabMailleRealise[9]),FloatToStr(TabMailleRealise[10]),
                                         FloatToStr(TabMailleRealise[11]),FloatToStr(TabMailleRealise[12]),
                                         FloatToStr(TabMailleRealise[13]),FloatToStr(TabMailleRealise[14]),
                                         FloatToStr(TabMailleRealise[15]),FloatToStr(TabMailleRealise[16]),
                                         FloatToStr(TabMailleRealise[17]),FloatToStr(TabMailleRealise[18]),
                                         FloatToStr(TabMailleRealise[19]),FloatToStr(TabMailleRealise[20]),
                                         FloatToStr(TabMailleRealise[21]),FloatToStr(TabMailleRealise[22]),
                                         FloatToStr(TabMailleRealise[23]),FloatToStr(TabMailleRealise[24]),
                                         FloatToStr(TabMailleRealise[25]),FloatToStr(TabMailleRealise[26]),
                                         FloatToStr(TabMailleRealise[27]),FloatToStr(TabMailleRealise[28]),
                                         FloatToStr(TabMailleRealise[29]),FloatToStr(TabMailleRealise[30]),
                                         FloatToStr(TabMailleRealise[31]),FloatToStr(TabMailleRealise[32]),
                                         FloatToStr(TabMailleRealise[33]),FloatToStr(TabMailleRealise[34]),
                                         FloatToStr(TabMailleRealise[35]),FloatToStr(TabMailleRealise[36]),
                                         FloatToStr(TabMailleRealise[37]),FloatToStr(TabMailleRealise[38]),
                                         FloatToStr(TabMailleRealise[39]),FloatToStr(TabMailleRealise[40]),
                                         FloatToStr(TabMailleRealise[41]),FloatToStr(TabMailleRealise[42]),
                                         FloatToStr(TabMailleRealise[43]),FloatToStr(TabMailleRealise[44]),
                                         FloatToStr(TabMailleRealise[45]),FloatToStr(TabMailleRealise[46]),
                                         FloatToStr(TabMailleRealise[47]),FloatToStr(TabMailleRealise[48]),
                                         FloatToStr(TabMailleRealise[49]),FloatToStr(TabMailleRealise[50]),
                                         FloatToStr(TabMailleRealise[51]),FloatToStr(TabMailleRealise[52]),
                                         FloatToStr(TabMailleRealise[53])]);
                //CEGID-CCMX le 17/11/2006 DEBUT
                if (sTypeDA = 'TYPEDA') then
                begin
                  dm_creaEnreg(idm_tmpBPB,['4-PRESQUE REALISE',TabCodeAxeR[1],TabCodeAxeR[2],
                                           TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                           TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                           TabCodeAxeR[9],TabCodeAxeR[10],
                                           TabLibValAxe[1],TabLibValAxe[2],
                                           TabLibValAxe[3],TabLibValAxe[4],
                                           TabLibValAxe[5],TabLibValAxe[6],
                                           TabLibValAxe[7],TabLibValAxe[8],
                                           TabLibValAxe[9],TabLibValAxe[10],
                                           FloatToStr(TabMaillePReal[1]),FloatToStr(TabMaillePReal[2]),
                                           FloatToStr(TabMaillePReal[3]),FloatToStr(TabMaillePReal[4]),
                                           FloatToStr(TabMaillePReal[5]),FloatToStr(TabMaillePReal[6]),
                                           FloatToStr(TabMaillePReal[7]),FloatToStr(TabMaillePReal[8]),
                                           FloatToStr(TabMaillePReal[9]),FloatToStr(TabMaillePReal[10]),
                                           FloatToStr(TabMaillePReal[11]),FloatToStr(TabMaillePReal[12]),
                                           FloatToStr(TabMaillePReal[13]),FloatToStr(TabMaillePReal[14]),
                                           FloatToStr(TabMaillePReal[15]),FloatToStr(TabMaillePReal[16]),
                                           FloatToStr(TabMaillePReal[17]),FloatToStr(TabMaillePReal[18]),
                                           FloatToStr(TabMaillePReal[19]),FloatToStr(TabMaillePReal[20]),
                                           FloatToStr(TabMaillePReal[21]),FloatToStr(TabMaillePReal[22]),
                                           FloatToStr(TabMaillePReal[23]),FloatToStr(TabMaillePReal[24]),
                                           FloatToStr(TabMaillePReal[25]),FloatToStr(TabMaillePReal[26]),
                                           FloatToStr(TabMaillePReal[27]),FloatToStr(TabMaillePReal[28]),
                                           FloatToStr(TabMaillePReal[29]),FloatToStr(TabMaillePReal[30]),
                                           FloatToStr(TabMaillePReal[31]),FloatToStr(TabMaillePReal[32]),
                                           FloatToStr(TabMaillePReal[33]),FloatToStr(TabMaillePReal[34]),
                                           FloatToStr(TabMaillePReal[35]),FloatToStr(TabMaillePReal[36]),
                                           FloatToStr(TabMaillePReal[37]),FloatToStr(TabMaillePReal[38]),
                                           FloatToStr(TabMaillePReal[39]),FloatToStr(TabMaillePReal[40]),
                                           FloatToStr(TabMaillePReal[41]),FloatToStr(TabMaillePReal[42]),
                                           FloatToStr(TabMaillePReal[43]),FloatToStr(TabMaillePReal[44]),
                                           FloatToStr(TabMaillePReal[45]),FloatToStr(TabMaillePReal[46]),
                                           FloatToStr(TabMaillePReal[47]),FloatToStr(TabMaillePReal[48]),
                                           FloatToStr(TabMaillePReal[49]),FloatToStr(TabMaillePReal[50]),
                                           FloatToStr(TabMaillePReal[51]),FloatToStr(TabMaillePReal[52]),
                                           FloatToStr(TabMaillePReal[53])]);
                end;
                //CEGID-CCMX le 17/11/2006 FIN
                if BPOkOrli then
                begin //-----------------> ORLI
                  dm_creaEnreg(idm_tmpBPB,['4-QTEHISTO',TabCodeAxeR[1],TabCodeAxeR[2],
                                           TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                           TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                           TabCodeAxeR[9],TabCodeAxeR[10],
                                           TabLibValAxe[1],TabLibValAxe[2],
                                           TabLibValAxe[3],TabLibValAxe[4],
                                           TabLibValAxe[5],TabLibValAxe[6],
                                           TabLibValAxe[7],TabLibValAxe[8],
                                           TabLibValAxe[9],TabLibValAxe[10],
                                           FloatToStr(TabMailleQteHisto[1]),FloatToStr(TabMailleQteHisto[2]),
                                           FloatToStr(TabMailleQteHisto[3]),FloatToStr(TabMailleQteHisto[4]),
                                           FloatToStr(TabMailleQteHisto[5]),FloatToStr(TabMailleQteHisto[6]),
                                           FloatToStr(TabMailleQteHisto[7]),FloatToStr(TabMailleQteHisto[8]),
                                           FloatToStr(TabMailleQteHisto[9]),FloatToStr(TabMailleQteHisto[10]),
                                           FloatToStr(TabMailleQteHisto[11]),FloatToStr(TabMailleQteHisto[12]),
                                           FloatToStr(TabMailleQteHisto[13]),FloatToStr(TabMailleQteHisto[14]),
                                           FloatToStr(TabMailleQteHisto[15]),FloatToStr(TabMailleQteHisto[16]),
                                           FloatToStr(TabMailleQteHisto[17]),FloatToStr(TabMailleQteHisto[18]),
                                           FloatToStr(TabMailleQteHisto[19]),FloatToStr(TabMailleQteHisto[20]),
                                           FloatToStr(TabMailleQteHisto[21]),FloatToStr(TabMailleQteHisto[22]),
                                           FloatToStr(TabMailleQteHisto[23]),FloatToStr(TabMailleQteHisto[24]),
                                           FloatToStr(TabMailleQteHisto[25]),FloatToStr(TabMailleQteHisto[26]),
                                           FloatToStr(TabMailleQteHisto[27]),FloatToStr(TabMailleQteHisto[28]),
                                           FloatToStr(TabMailleQteHisto[29]),FloatToStr(TabMailleQteHisto[30]),
                                           FloatToStr(TabMailleQteHisto[31]),FloatToStr(TabMailleQteHisto[32]),
                                           FloatToStr(TabMailleQteHisto[33]),FloatToStr(TabMailleQteHisto[34]),
                                           FloatToStr(TabMailleQteHisto[35]),FloatToStr(TabMailleQteHisto[36]),
                                           FloatToStr(TabMailleQteHisto[37]),FloatToStr(TabMailleQteHisto[38]),
                                           FloatToStr(TabMailleQteHisto[39]),FloatToStr(TabMailleQteHisto[40]),
                                           FloatToStr(TabMailleQteHisto[41]),FloatToStr(TabMailleQteHisto[42]),
                                           FloatToStr(TabMailleQteHisto[43]),FloatToStr(TabMailleQteHisto[44]),
                                           FloatToStr(TabMailleQteHisto[45]),FloatToStr(TabMailleQteHisto[46]),
                                           FloatToStr(TabMailleQteHisto[47]),FloatToStr(TabMailleQteHisto[48]),
                                           FloatToStr(TabMailleQteHisto[49]),FloatToStr(TabMailleQteHisto[50]),
                                           FloatToStr(TabMailleQteHisto[51]),FloatToStr(TabMailleQteHisto[52]),
                                           FloatToStr(TabMailleQteHisto[53])]);
                  dm_creaEnreg(idm_tmpBPB,['5-QTEPREVU',TabCodeAxeR[1],TabCodeAxeR[2],
                                           TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                           TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                           TabCodeAxeR[9],TabCodeAxeR[10],
                                           TabLibValAxe[1],TabLibValAxe[2],
                                           TabLibValAxe[3],TabLibValAxe[4],
                                           TabLibValAxe[5],TabLibValAxe[6],
                                           TabLibValAxe[7],TabLibValAxe[8],
                                           TabLibValAxe[9],TabLibValAxe[10],
                                           FloatToStr(TabMailleQtePrevu[1]),FloatToStr(TabMailleQtePrevu[2]),
                                           FloatToStr(TabMailleQtePrevu[3]),FloatToStr(TabMailleQtePrevu[4]),
                                           FloatToStr(TabMailleQtePrevu[5]),FloatToStr(TabMailleQtePrevu[6]),
                                           FloatToStr(TabMailleQtePrevu[7]),FloatToStr(TabMailleQtePrevu[8]),
                                           FloatToStr(TabMailleQtePrevu[9]),FloatToStr(TabMailleQtePrevu[10]),
                                           FloatToStr(TabMailleQtePrevu[11]),FloatToStr(TabMailleQtePrevu[12]),
                                           FloatToStr(TabMailleQtePrevu[13]),FloatToStr(TabMailleQtePrevu[14]),
                                           FloatToStr(TabMailleQtePrevu[15]),FloatToStr(TabMailleQtePrevu[16]),
                                           FloatToStr(TabMailleQtePrevu[17]),FloatToStr(TabMailleQtePrevu[18]),
                                           FloatToStr(TabMailleQtePrevu[19]),FloatToStr(TabMailleQtePrevu[20]),
                                           FloatToStr(TabMailleQtePrevu[21]),FloatToStr(TabMailleQtePrevu[22]),
                                           FloatToStr(TabMailleQtePrevu[23]),FloatToStr(TabMailleQtePrevu[24]),
                                           FloatToStr(TabMailleQtePrevu[25]),FloatToStr(TabMailleQtePrevu[26]),
                                           FloatToStr(TabMailleQtePrevu[27]),FloatToStr(TabMailleQtePrevu[28]),
                                           FloatToStr(TabMailleQtePrevu[29]),FloatToStr(TabMailleQtePrevu[30]),
                                           FloatToStr(TabMailleQtePrevu[31]),FloatToStr(TabMailleQtePrevu[32]),
                                           FloatToStr(TabMailleQtePrevu[33]),FloatToStr(TabMailleQtePrevu[34]),
                                           FloatToStr(TabMailleQtePrevu[35]),FloatToStr(TabMailleQtePrevu[36]),
                                           FloatToStr(TabMailleQtePrevu[37]),FloatToStr(TabMailleQtePrevu[38]),
                                           FloatToStr(TabMailleQtePrevu[39]),FloatToStr(TabMailleQtePrevu[40]),
                                           FloatToStr(TabMailleQtePrevu[41]),FloatToStr(TabMailleQtePrevu[42]),
                                           FloatToStr(TabMailleQtePrevu[43]),FloatToStr(TabMailleQtePrevu[44]),
                                           FloatToStr(TabMailleQtePrevu[45]),FloatToStr(TabMailleQtePrevu[46]),
                                           FloatToStr(TabMailleQtePrevu[47]),FloatToStr(TabMailleQtePrevu[48]),
                                           FloatToStr(TabMailleQtePrevu[49]),FloatToStr(TabMailleQtePrevu[50]),
                                           FloatToStr(TabMailleQtePrevu[51]),FloatToStr(TabMailleQtePrevu[52]),
                                           FloatToStr(TabMailleQtePrevu[53])]);
                  dm_creaEnreg(idm_tmpBPB,['6-QTEREALISE',TabCodeAxeR[1],TabCodeAxeR[2],
                                           TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                           TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                           TabCodeAxeR[9],TabCodeAxeR[10],
                                           TabLibValAxe[1],TabLibValAxe[2],
                                           TabLibValAxe[3],TabLibValAxe[4],
                                           TabLibValAxe[5],TabLibValAxe[6],
                                           TabLibValAxe[7],TabLibValAxe[8],
                                           TabLibValAxe[9],TabLibValAxe[10],
                                           FloatToStr(TabMailleQteRealise[1]),FloatToStr(TabMailleQteRealise[2]),
                                           FloatToStr(TabMailleQteRealise[3]),FloatToStr(TabMailleQteRealise[4]),
                                           FloatToStr(TabMailleQteRealise[5]),FloatToStr(TabMailleQteRealise[6]),
                                           FloatToStr(TabMailleQteRealise[7]),FloatToStr(TabMailleQteRealise[8]),
                                           FloatToStr(TabMailleQteRealise[9]),FloatToStr(TabMailleQteRealise[10]),
                                           FloatToStr(TabMailleQteRealise[11]),FloatToStr(TabMailleQteRealise[12]),
                                           FloatToStr(TabMailleQteRealise[13]),FloatToStr(TabMailleQteRealise[14]),
                                           FloatToStr(TabMailleQteRealise[15]),FloatToStr(TabMailleQteRealise[16]),
                                           FloatToStr(TabMailleQteRealise[17]),FloatToStr(TabMailleQteRealise[18]),
                                           FloatToStr(TabMailleQteRealise[19]),FloatToStr(TabMailleQteRealise[20]),
                                           FloatToStr(TabMailleQteRealise[21]),FloatToStr(TabMailleQteRealise[22]),
                                           FloatToStr(TabMailleQteRealise[23]),FloatToStr(TabMailleQteRealise[24]),
                                           FloatToStr(TabMailleQteRealise[25]),FloatToStr(TabMailleQteRealise[26]),
                                           FloatToStr(TabMailleQteRealise[27]),FloatToStr(TabMailleQteRealise[28]),
                                           FloatToStr(TabMailleQteRealise[29]),FloatToStr(TabMailleQteRealise[30]),
                                           FloatToStr(TabMailleQteRealise[31]),FloatToStr(TabMailleQteRealise[32]),
                                           FloatToStr(TabMailleQteRealise[33]),FloatToStr(TabMailleQteRealise[34]),
                                           FloatToStr(TabMailleQteRealise[35]),FloatToStr(TabMailleQteRealise[36]),
                                           FloatToStr(TabMailleQteRealise[37]),FloatToStr(TabMailleQteRealise[38]),
                                           FloatToStr(TabMailleQteRealise[39]),FloatToStr(TabMailleQteRealise[40]),
                                           FloatToStr(TabMailleQteRealise[41]),FloatToStr(TabMailleQteRealise[42]),
                                           FloatToStr(TabMailleQteRealise[43]),FloatToStr(TabMailleQteRealise[44]),
                                           FloatToStr(TabMailleQteRealise[45]),FloatToStr(TabMailleQteRealise[46]),
                                           FloatToStr(TabMailleQteRealise[47]),FloatToStr(TabMailleQteRealise[48]),
                                           FloatToStr(TabMailleQteRealise[49]),FloatToStr(TabMailleQteRealise[50]),
                                           FloatToStr(TabMailleQteRealise[51]),FloatToStr(TabMailleQteRealise[52]),
                                           FloatToStr(TabMailleQteRealise[53])]);
                end; //ORLI <-----------------
              end;
            end;
      3 : begin
            for i := 1 to ValAff do
            begin
              Enreg := false;
              if UpperCase(Q.fields[(nivMax*2)+1].asString) = UpperCase('Théorique') then
              begin
                if dm_trouveenr(idm_tmpBPB,[TabLib[i]+TraduireMemoire(' [Théorique]'),TabCodeAxeR[1],TabCodeAxeR[2],
                                TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                TabCodeAxeR[9],TabCodeAxeR[10]],enrBP)=0
                then
                begin
                  dm_modifchcour(idm_tmpBPB,diffdate,FloatToStr(valeur(enrBP.ch(diffdate))+(TabHisto[i-1])));
                  Enreg := true;
                end;
              end;

              if UpperCase(Q.fields[(nivMax*2)+1].asString) = UpperCase('Budget N') then
              begin
                if dm_trouveenr(idm_tmpBPB,[TabLib[i]+TraduireMemoire(' [Budget N]'),TabCodeAxeR[1],TabCodeAxeR[2],
                                TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                TabCodeAxeR[9],TabCodeAxeR[10]],enrBP)=0
                then
                begin
                  dm_modifchcour(idm_tmpBPB,diffdate,FloatToStr(valeur(enrBP.ch(diffdate))+(TabPrevu[i-1])));
                  Enreg := true;
                end ;
              end;

              if UpperCase(Q.fields[(nivMax*2)+1].asString) = UpperCase('Réalisé N') then
              begin
                if dm_trouveenr(idm_tmpBPB,[TabLib[i]+TraduireMemoire(' [Réalisé N]'),TabCodeAxeR[1],TabCodeAxeR[2],
                                 TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                 TabCodeAxeR[9],TabCodeAxeR[10]],enrBP)=0
                then
                begin
                  dm_modifchcour(idm_tmpBPB,diffdate,FloatToStr(valeur(enrBP.ch(diffdate))+(TabRealiseN[i-1])));
                  Enreg := true;
                end;
              end;

              if UpperCase(Q.fields[(nivMax*2)+1].asString) = UpperCase('Réalisé N-1') then
              begin
                if dm_trouveenr(idm_tmpBPB,[TabLib[i]+TraduireMemoire(' [Réalisé N-1]'),TabCodeAxeR[1],TabCodeAxeR[2],
                                 TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                 TabCodeAxeR[9],TabCodeAxeR[10]],enrBP)=0
                then
                begin
                  dm_modifchcour(idm_tmpBPB,diffdate,FloatToStr(valeur(enrBP.ch(diffdate))+(TabRealiseNm[i-1])));
                  Enreg := true;
                end;
              end;

              if Enreg = false then
              begin
                for k:=1 to 53 do
                begin
                  TabMailleHisto[k]:=0;
                  TabMaillePrevu[k]:=0;
                  tabMailleRealise[k]:=0;
                  tabMailleRealiseN[k]:=0;
                  if k=diffdate-21 then
                  begin
                    TabMailleHisto[k]:=TabHisto[i-1];
                    TabMaillePrevu[k]:=TabPrevu[i-1];
                    tabMailleRealise[k]:=TabRealiseN[i-1];
                    tabMailleRealiseN[k]:=TabRealiseNm[i-1];
                  end;
                end;
                if UpperCase(Q.fields[(nivMax*2)+1].asString) = UpperCase('Théorique') then
                dm_creaEnreg(idm_tmpBPB,[Base,TabLib[i]+TraduireMemoire(' [Théorique]'),TabCodeAxeR[1],TabCodeAxeR[2],
                                         TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                         TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                         TabCodeAxeR[9],TabCodeAxeR[10],
                                         TabLibValAxe[1],TabLibValAxe[2],
                                         TabLibValAxe[3],TabLibValAxe[4],
                                         TabLibValAxe[5],TabLibValAxe[6],
                                         TabLibValAxe[7],TabLibValAxe[8],
                                         TabLibValAxe[9],TabLibValAxe[10],
                                         FloatToStr(TabMailleHisto[1]),FloatToStr(TabMailleHisto[2]),
                                         FloatToStr(TabMailleHisto[3]),FloatToStr(TabMailleHisto[4]),
                                         FloatToStr(TabMailleHisto[5]),FloatToStr(TabMailleHisto[6]),
                                         FloatToStr(TabMailleHisto[7]),FloatToStr(TabMailleHisto[8]),
                                         FloatToStr(TabMailleHisto[9]),FloatToStr(TabMailleHisto[10]),
                                         FloatToStr(TabMailleHisto[11]),FloatToStr(TabMailleHisto[12]),
                                         FloatToStr(TabMailleHisto[13]),FloatToStr(TabMailleHisto[14]),
                                         FloatToStr(TabMailleHisto[15]),FloatToStr(TabMailleHisto[16]),
                                         FloatToStr(TabMailleHisto[17]),FloatToStr(TabMailleHisto[18]),
                                         FloatToStr(TabMailleHisto[19]),FloatToStr(TabMailleHisto[20]),
                                         FloatToStr(TabMailleHisto[21]),FloatToStr(TabMailleHisto[22]),
                                         FloatToStr(TabMailleHisto[23]),FloatToStr(TabMailleHisto[24]),
                                         FloatToStr(TabMailleHisto[25]),FloatToStr(TabMailleHisto[26]),
                                         FloatToStr(TabMailleHisto[27]),FloatToStr(TabMailleHisto[28]),
                                         FloatToStr(TabMailleHisto[29]),FloatToStr(TabMailleHisto[30]),
                                         FloatToStr(TabMailleHisto[31]),FloatToStr(TabMailleHisto[32]),
                                         FloatToStr(TabMailleHisto[33]),FloatToStr(TabMailleHisto[34]),
                                         FloatToStr(TabMailleHisto[35]),FloatToStr(TabMailleHisto[36]),
                                         FloatToStr(TabMailleHisto[37]),FloatToStr(TabMailleHisto[38]),
                                         FloatToStr(TabMailleHisto[39]),FloatToStr(TabMailleHisto[40]),
                                         FloatToStr(TabMailleHisto[41]),FloatToStr(TabMailleHisto[42]),
                                         FloatToStr(TabMailleHisto[43]),FloatToStr(TabMailleHisto[44]),
                                         FloatToStr(TabMailleHisto[45]),FloatToStr(TabMailleHisto[46]),
                                         FloatToStr(TabMailleHisto[47]),FloatToStr(TabMailleHisto[48]),
                                         FloatToStr(TabMailleHisto[49]),FloatToStr(TabMailleHisto[50]),
                                         FloatToStr(TabMailleHisto[51]),FloatToStr(TabMailleHisto[52]),
                                         FloatToStr(TabMailleHisto[53])]);
                if UpperCase(Q.fields[(nivMax*2)+1].asString) = UpperCase('Budget N') then
                dm_creaEnreg(idm_tmpBPB,[Base,TabLib[i]+TraduireMemoire(' [Budget N]'),TabCodeAxeR[1],TabCodeAxeR[2],
                                         TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                         TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                         TabCodeAxeR[9],TabCodeAxeR[10],
                                         TabLibValAxe[1],TabLibValAxe[2],
                                         TabLibValAxe[3],TabLibValAxe[4],
                                         TabLibValAxe[5],TabLibValAxe[6],
                                         TabLibValAxe[7],TabLibValAxe[8],
                                         TabLibValAxe[9],TabLibValAxe[10],
                                         FloatToStr(TabMaillePrevu[1]),FloatToStr(TabMaillePrevu[2]),
                                         FloatToStr(TabMaillePrevu[3]),FloatToStr(TabMaillePrevu[4]),
                                         FloatToStr(TabMaillePrevu[5]),FloatToStr(TabMaillePrevu[6]),
                                         FloatToStr(TabMaillePrevu[7]),FloatToStr(TabMaillePrevu[8]),
                                         FloatToStr(TabMaillePrevu[9]),FloatToStr(TabMaillePrevu[10]),
                                         FloatToStr(TabMaillePrevu[11]),FloatToStr(TabMaillePrevu[12]),
                                         FloatToStr(TabMaillePrevu[13]),FloatToStr(TabMaillePrevu[14]),
                                         FloatToStr(TabMaillePrevu[15]),FloatToStr(TabMaillePrevu[16]),
                                         FloatToStr(TabMaillePrevu[17]),FloatToStr(TabMaillePrevu[18]),
                                         FloatToStr(TabMaillePrevu[19]),FloatToStr(TabMaillePrevu[20]),
                                         FloatToStr(TabMaillePrevu[21]),FloatToStr(TabMaillePrevu[22]),
                                         FloatToStr(TabMaillePrevu[23]),FloatToStr(TabMaillePrevu[24]),
                                         FloatToStr(TabMaillePrevu[25]),FloatToStr(TabMaillePrevu[26]),
                                         FloatToStr(TabMaillePrevu[27]),FloatToStr(TabMaillePrevu[28]),
                                         FloatToStr(TabMaillePrevu[29]),FloatToStr(TabMaillePrevu[30]),
                                         FloatToStr(TabMaillePrevu[31]),FloatToStr(TabMaillePrevu[32]),
                                         FloatToStr(TabMaillePrevu[33]),FloatToStr(TabMaillePrevu[34]),
                                         FloatToStr(TabMaillePrevu[35]),FloatToStr(TabMaillePrevu[36]),
                                         FloatToStr(TabMaillePrevu[37]),FloatToStr(TabMaillePrevu[38]),
                                         FloatToStr(TabMaillePrevu[39]),FloatToStr(TabMaillePrevu[40]),
                                         FloatToStr(TabMaillePrevu[41]),FloatToStr(TabMaillePrevu[42]),
                                         FloatToStr(TabMaillePrevu[43]),FloatToStr(TabMaillePrevu[44]),
                                         FloatToStr(TabMaillePrevu[45]),FloatToStr(TabMaillePrevu[46]),
                                         FloatToStr(TabMaillePrevu[47]),FloatToStr(TabMaillePrevu[48]),
                                         FloatToStr(TabMaillePrevu[49]),FloatToStr(TabMaillePrevu[50]),
                                         FloatToStr(TabMaillePrevu[51]),FloatToStr(TabMaillePrevu[52]),
                                         FloatToStr(TabMaillePrevu[53])]);
                if UpperCase(Q.fields[(nivMax*2)+1].asString) = UpperCase('Réalisé N') then
                dm_creaEnreg(idm_tmpBPB,[Base,TabLib[i]+TraduireMemoire(' [Réalisé N]'),TabCodeAxeR[1],TabCodeAxeR[2],
                                         TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                         TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                         TabCodeAxeR[9],TabCodeAxeR[10],
                                         TabLibValAxe[1],TabLibValAxe[2],
                                         TabLibValAxe[3],TabLibValAxe[4],
                                         TabLibValAxe[5],TabLibValAxe[6],
                                         TabLibValAxe[7],TabLibValAxe[8],
                                         TabLibValAxe[9],TabLibValAxe[10],
                                         FloatToStr(tabMailleRealise[1]),FloatToStr(tabMailleRealise[2]),
                                         FloatToStr(tabMailleRealise[3]),FloatToStr(tabMailleRealise[4]),
                                         FloatToStr(tabMailleRealise[5]),FloatToStr(tabMailleRealise[6]),
                                         FloatToStr(tabMailleRealise[7]),FloatToStr(tabMailleRealise[8]),
                                         FloatToStr(tabMailleRealise[9]),FloatToStr(tabMailleRealise[10]),
                                         FloatToStr(tabMailleRealise[11]),FloatToStr(tabMailleRealise[12]),
                                         FloatToStr(tabMailleRealise[13]),FloatToStr(tabMailleRealise[14]),
                                         FloatToStr(tabMailleRealise[15]),FloatToStr(tabMailleRealise[16]),
                                         FloatToStr(tabMailleRealise[17]),FloatToStr(tabMailleRealise[18]),
                                         FloatToStr(tabMailleRealise[19]),FloatToStr(tabMailleRealise[20]),
                                         FloatToStr(tabMailleRealise[21]),FloatToStr(tabMailleRealise[22]),
                                         FloatToStr(tabMailleRealise[23]),FloatToStr(tabMailleRealise[24]),
                                         FloatToStr(tabMailleRealise[25]),FloatToStr(tabMailleRealise[26]),
                                         FloatToStr(tabMailleRealise[27]),FloatToStr(tabMailleRealise[28]),
                                         FloatToStr(tabMailleRealise[29]),FloatToStr(tabMailleRealise[30]),
                                         FloatToStr(tabMailleRealise[31]),FloatToStr(tabMailleRealise[32]),
                                         FloatToStr(tabMailleRealise[33]),FloatToStr(tabMailleRealise[34]),
                                         FloatToStr(tabMailleRealise[35]),FloatToStr(tabMailleRealise[36]),
                                         FloatToStr(tabMailleRealise[37]),FloatToStr(tabMailleRealise[38]),
                                         FloatToStr(tabMailleRealise[39]),FloatToStr(tabMailleRealise[40]),
                                         FloatToStr(tabMailleRealise[41]),FloatToStr(tabMailleRealise[42]),
                                         FloatToStr(tabMailleRealise[43]),FloatToStr(tabMailleRealise[44]),
                                         FloatToStr(tabMailleRealise[45]),FloatToStr(tabMailleRealise[46]),
                                         FloatToStr(tabMailleRealise[47]),FloatToStr(tabMailleRealise[48]),
                                         FloatToStr(tabMailleRealise[49]),FloatToStr(tabMailleRealise[50]),
                                         FloatToStr(tabMailleRealise[51]),FloatToStr(tabMailleRealise[52]),
                                         FloatToStr(tabMailleRealise[53])]);
                if UpperCase(Q.fields[(nivMax*2)+1].asString) = UpperCase('Réalisé N-1') then
                dm_creaEnreg(idm_tmpBPB,[Base,TabLib[i]+TraduireMemoire(' [Réalisé N-1]'),TabCodeAxeR[1],TabCodeAxeR[2],
                                         TabCodeAxeR[3],TabCodeAxeR[4],TabCodeAxeR[5],
                                         TabCodeAxeR[6],TabCodeAxeR[7],TabCodeAxeR[8],
                                         TabCodeAxeR[9],TabCodeAxeR[10],
                                         TabLibValAxe[1],TabLibValAxe[2],
                                         TabLibValAxe[3],TabLibValAxe[4],
                                         TabLibValAxe[5],TabLibValAxe[6],
                                         TabLibValAxe[7],TabLibValAxe[8],
                                         TabLibValAxe[9],TabLibValAxe[10],
                                         FloatToStr(tabMailleRealiseN[1]),FloatToStr(tabMailleRealiseN[2]),
                                         FloatToStr(tabMailleRealiseN[3]),FloatToStr(tabMailleRealiseN[4]),
                                         FloatToStr(tabMailleRealiseN[5]),FloatToStr(tabMailleRealiseN[6]),
                                         FloatToStr(tabMailleRealiseN[7]),FloatToStr(tabMailleRealiseN[8]),
                                         FloatToStr(tabMailleRealiseN[9]),FloatToStr(tabMailleRealiseN[10]),
                                         FloatToStr(tabMailleRealiseN[11]),FloatToStr(tabMailleRealiseN[12]),
                                         FloatToStr(tabMailleRealiseN[13]),FloatToStr(tabMailleRealiseN[14]),
                                         FloatToStr(tabMailleRealiseN[15]),FloatToStr(tabMailleRealiseN[16]),
                                         FloatToStr(tabMailleRealiseN[17]),FloatToStr(tabMailleRealiseN[18]),
                                         FloatToStr(tabMailleRealiseN[19]),FloatToStr(tabMailleRealiseN[20]),
                                         FloatToStr(tabMailleRealiseN[21]),FloatToStr(tabMailleRealiseN[22]),
                                         FloatToStr(tabMailleRealiseN[23]),FloatToStr(tabMailleRealiseN[24]),
                                         FloatToStr(tabMailleRealiseN[25]),FloatToStr(tabMailleRealiseN[26]),
                                         FloatToStr(tabMailleRealiseN[27]),FloatToStr(tabMailleRealiseN[28]),
                                         FloatToStr(tabMailleRealiseN[29]),FloatToStr(tabMailleRealiseN[30]),
                                         FloatToStr(tabMailleRealiseN[31]),FloatToStr(tabMailleRealiseN[32]),
                                         FloatToStr(tabMailleRealiseN[33]),FloatToStr(tabMailleRealiseN[34]),
                                         FloatToStr(tabMailleRealiseN[35]),FloatToStr(tabMailleRealiseN[36]),
                                         FloatToStr(tabMailleRealiseN[37]),FloatToStr(tabMailleRealiseN[38]),
                                         FloatToStr(tabMailleRealiseN[39]),FloatToStr(tabMailleRealiseN[40]),
                                         FloatToStr(tabMailleRealiseN[41]),FloatToStr(tabMailleRealiseN[42]),
                                         FloatToStr(tabMailleRealiseN[43]),FloatToStr(tabMailleRealiseN[44]),
                                         FloatToStr(tabMailleRealiseN[45]),FloatToStr(tabMailleRealiseN[46]),
                                         FloatToStr(tabMailleRealiseN[47]),FloatToStr(tabMailleRealiseN[48]),
                                         FloatToStr(tabMailleRealiseN[49]),FloatToStr(tabMailleRealiseN[50]),
                                         FloatToStr(tabMailleRealiseN[51]),FloatToStr(tabMailleRealiseN[52]),
                                         FloatToStr(tabMailleRealiseN[53])]);
              end
            end
          end

      end// CASE

    end;
    Q.next
  end;
  ferme(Q);
  FiniMoveProgressForm;
  InProgress := false;
end;


procedure TOF_QUFSBPBUDGETR3.OnUpdate ;
var enrBP:TEnreg;
    Pseq:TdmPseq;
    MailleS,ChpS:hString;
    i,NbChp,MultiSocWarning,rep:integer;
    TabMaille,TabChp:array of String;
    TabChpVal:array of variant;
    TabMailleD:array of double;
    OkDelai,MultiSocError:boolean;
    _DateDebRef,_DateFinRef:TDateTime; // ne sert pas
    lstMsg,lstMsgState:TStringList;
    sSocietes:string;
begin
  Inherited ;

  MultiSocWarning:=0;
  MultiSocError:=false;
  blocageMenu(TFStat(ecran));
  //lit le code session et le vérifie
  //lit les dates pour l'intervalle de calcul
  //si le code session est incorrect ou non renseigné on ne fait rien
  //si les dates sont incorrectes on ne fait rien
  edtsession := GetControlText('EDTSESSION');

  if (ControleSession(edtsession)=false) then
  begin
    PGIINFO('Le code session doit être renseigné correctement.');
    deblocageMenu(TFStat(ecran));
    exit;
  end
  else if SessionObjInitDelai(edtsession)=false then
  begin
    PGIINFO('Cette session n''est pas initialisée.');
    deblocageMenu(TFStat(ecran));
    exit;
  end
  else
  begin
    sSocietes := GetBasesMS(THValComboBox(GetControl('MULTIDOSSIER')).Value, true);
    if sSocietes <> '' then
    begin
      //Contrôle des sessions
      lstMsg:=TStringList.Create;
      lstMsgState:=TStringList.Create;
      VerifCoherenceSessions(sSocietes,edtsession,lstMsg,lstMsgState);
      for i:=0 to lstMsg.Count-1 do
      begin
        if lstMsgState[i]='$000080FF' then MultiSocWarning:=MultiSocWarning+1;
        if lstMsgState[i]='clRed' then
        begin
          MultiSocError:=true;
          break;
        end;
      end;
      lstMsg.free;
      lstMsgState.free;

      if MultiSocError=true then
      begin
        PGIINFO('Le contôle multi sociétés a renvoyé des erreurs.');
        deblocageMenu(TFStat(ecran));
        exit;
      end
      else
      begin
        if MultiSocWarning>0 then
        begin
          rep:=HShowMessage('1;Attention;Le contôle multi sociétés a renvoyé des avertissements.'+#13#10+
                            ' Voulez-vous continuer ?;Q;YN;N;N;','','');
          if rep = mrNo then
          begin
            deblocageMenu(TFStat(ecran));
            exit;
          end
        end
      end
    end
  end;


  InitMoveProgressForm(nil, TraduireMemoire('Tableau de bord'),TraduireMemoire('Veuillez patienter'), NbDeMaille(edtsession) + NbMajCoeff(edtsession), True, True);
  InProgress := true;
  //recherche de la période courante et courante-1
  //à partir de la session entrée
  ChercheDateDDateFPeriode(edtsession,edtDateDeb,edtDateFin,_DateDebRef,_DateFinRef);
  if  (edtDateDeb=0) or (edtDateFin=0) then
  begin
    PGIINFO('Les dates pour la période courante doivent être renseignées au niveau de la session pour avoir ce tableau de bord');
    deblocageMenu(TFStat(ecran));
    exit;
  end;
  structure:=ChercheSessionStructure(edtSession);
  ChercheTabCodeAxeTabNumValAxe(edtSession,TabCodeAxe,TabNumValAxe);

  //cherche le nombre de niveau definit
  nivMax:=ChercheNivMax(edtsession);

  BPInitialise:=SessionBPInitialise(edtSession);

  if (BPInitialise='') and SessionDelai(edtSession) then BPInitialise:='4';
  if (BPInitialise='0') and (SessionInitPrev(edtSession)) then BPInitialise:='4';
  if SessionEclateeParDelai(edtSession) then BPInitialise:='4';
  if BPInitialise='5' then BPInitialise:='4';
  { if (BPInitialise='0') and (not OkSessionPermanent(edtSession))
      and (not OkSessionObjectif(edtSession))
   then BPInitialise:='4';
  }
  nbIntervalle:=1;
  //initialisation, suivant les coches, pour les colonnes mailles
  if CocheSemaine then
  begin
    NumSemDeb:= numSemaine(edtDateDeb,anDeb);
    //cherche nb semaine entre dates données
    nbIntervalle:=NbSemaineIntervalle(edtDateDeb,edtDateFin);
    if nbIntervalle>53 then nbIntervalle:=53;
    SetLength(TabMaille,nbIntervalle);
    for i:=1 to nbIntervalle do
    begin
      MailleS:=MailleS+';MAILLE'+IntToStr(i);
      TabMaille[i-1]:='MAILLE'+IntToStr(i);
    end;
  end;
  if CocheQuinzaine then
  begin
    Date_AnNumQuinz(edtDateDeb,anDebQ,moisDebQ,QuinzDeb,NumQuinzDeb);
    //cherche nb quinzaine entre dates données
    nbIntervalle:=NbQuinzaineIntervalle(edtDateDeb,edtDateFin);
    if nbIntervalle>53 then nbIntervalle:=53;
    SetLength(TabMaille,nbIntervalle);
    for i:=1 to nbIntervalle do
    begin
      MailleS:=MailleS+';MAILLE'+IntToStr(i);
      TabMaille[i-1]:='MAILLE'+IntToStr(i);
    end;
  end;
  if CocheMois then
  begin
    DecodeDate(edtDateDeb,anM,moisM,jourM);
    //cherche nb mois entre dates données
    nbIntervalle:=NbMoisIntervalle(edtDateDeb,edtDateFin);
    if nbIntervalle>53 then nbIntervalle:=53;
    SetLength(TabMaille,nbIntervalle);
    for i:=1 to nbIntervalle do
    begin
      MailleS:=MailleS+';MAILLE'+IntToStr(i);
      TabMaille[i-1]:='MAILLE'+IntToStr(i);
    end;
  end;

  OkDelai:=false;
  if (CocheSemaine) or (CocheQuinzaine) or (CocheMois) or
     (CocheMois445) then OkDelai:=true;
  if not OkDelai then
  begin
    nbIntervalle:=1;
    SetLength(TabMaille,nbIntervalle);
    TabMaille[0]:='MAILLE1';
    MailleS:=';MAILLE1';
  end;

  //dm
  dm_init(idm_tmpBPB, 0, ['A','A','A','A','A','A','A','A','A','A','A']
                       , true, [rc_BPAXE,rc_BPAXE,rc_BPAXE,rc_BPAXE,rc_BPAXE,
                                rc_BPAXE,rc_BPAXE,rc_BPAXE,rc_BPAXE,rc_BPAXE,
                                rc_BPAXE],DmBP);

  if SessionInitPrev(edtSession) then
  begin
    nivMax:=NivMax+1;
    TabCodeAxe[NivMax+1]:='ARTICLE';
    TabNumValAxe[NivMax+1]:='2';
    //niv coloris
    if GestionBPColoris then
    begin
      TabCodeAxe[nivMax+2]:='COLORIS';
      TabNumValAxe[nivMax+2]:='3';
      nivMax:=nivMax+1;
    end;
    //niv FS
    if GestionBPFS then
    begin
      TabCodeAxe[nivMax+2]:='FS';
      TabNumValAxe[nivMax+2]:='4';
      nivMax:=nivMax+1;
    end;
    //niv Magasin
    if GestionBPMagasin then
    begin
      TabCodeAxe[nivMax+2]:='MAGASIN';
      TabNumValAxe[nivMax+2]:='7';
      nivMax:=nivMax+1;
    end;

  end;

  //remplit le dm
  CalculColonnes(OkDelai);

  //création de la tob
  //on crée 52 mailles pour avoir 1 année
  killTob;

  if ContextBP = 3 then
  begin
    NbChp:=(2*NivMax)+2;
    SetLength(TabChp,NbChp+1);
    TabChp[0]:='BASE';
    TabChp[1]:='TYPE';
    ChpS:='BASE;TYPE;';
    for i:=2 to NivMax+1 do
    begin
      TabChp[i]:='AXE'+IntToStr(i-1);
      ChpS:=ChpS+'AXE'+IntToStr(i-1)+';';
    end;
    for i:=NivMax+2 to NbChp-1 do
    begin
      TabChp[i]:='LIB'+IntToStr(i-NivMax-1);
      ChpS:=ChpS+'LIB'+IntToStr(i-NivMax-1)+';';
    end;

    SetLength(TabChpVal,NbChp);
    QTBP:=TQTob.create(TabChp,TabMaille);
    //remplit tob
    if dm_pseqAll(idm_tmpBPB, Pseq) = 0 then
    while dm_pseqlect(Pseq, enrBP) = 0 do
    begin
      SetLength(TabMailleD,nbIntervalle);
      for i:=0 to nbIntervalle-1 do TabMailleD[i]:=VALEUR(enrBP.ch(TMPBPB_MAILLE1+i));
      if sSocietes='' then TabChpVal[0]:=V_PGI.SchemaName
      else TabChpVal[0]:=enrBP.ch(TMPBPB_TYPE);
      TabChpVal[1]:=enrBP.ch(TMPBPB_TYPE+1);
      for i:=2 to NivMax+1 do TabChpVal[i]:=enrBP.ch(TMPBPB_AXE1+i-1);
      for i:=NivMax+2 to NbChp-1 do TabChpVal[i]:=enrBP.ch(TMPBPB_LIB1+i-NivMax-1);

      QTBP.addValeur(TabChpVal,TabMailleD);
    end;
  end
  else
  begin
    NbChp:=(2*NivMax)+1;
    SetLength(TabChp,NbChp+1);
    TabChp[0]:='TYPE';
    ChpS:='TYPE;';
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

    SetLength(TabChpVal,NbChp);
    QTBP:=TQTob.create(TabChp,TabMaille);
    //remplit tob
    if dm_pseqAll(idm_tmpBPB, Pseq) = 0 then
    while dm_pseqlect(Pseq, enrBP) = 0 do
    begin
      SetLength(TabMailleD,nbIntervalle);
      for i:=0 to nbIntervalle-1 do TabMailleD[i]:=VALEUR(enrBP.ch(TMPBPB_MAILLE1+i));
      TabChpVal[0]:=enrBP.ch(TMPBPB_TYPE);
      for i:=1 to NivMax do TabChpVal[i]:=enrBP.ch(TMPBPB_AXE1+i-1);
      for i:=NivMax+1 to NivMax*2 do TabChpVal[i]:=enrBP.ch(TMPBPB_LIB1+i-NivMax-1);

      QTBP.addValeur(TabChpVal,TabMailleD);
    end;
  end;

  LibereUnDm(idm_tmpBPB);

  TFStat(Ecran).LaTOB := QTBP.latob; //TOB
  TFStat(Ecran).ColNames := ChpS+MailleS; //Colonnes

  deblocageMenu(TFStat(ecran));
end ;


procedure TOF_QUFSBPBUDGETR3.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_QUFSBPBUDGETR3.OnArgument (S : String ) ;
begin
  Inherited ;
  TTobViewer(getcontrol('TV')).OnDblClick := DblClick;
  { EVI / Trap de la touche ESC : éviter de sortir pendant la génération }
  Ecran.OnKeyDown := FormKeyDown;
  InProgress := false;
  BCherche := GetControl('BCherche') as TToolbarButton97;
  if v_pgi.RequeteMultiDossier.count = 0 then
  begin
    THBevel(GetControl('BVMULTIDOSSIER')).Hide;
    THLabel(GetControl('TMULTIDOSSIER')).Hide;
    THValComboBox(GetControl('MULTIDOSSIER')).Hide;
  end;
end ;

procedure TOF_QUFSBPBUDGETR3.OnClose ;
begin
  Inherited ;
  killTob;
end ;

procedure TOF_QUFSBPBUDGETR3.OnDisplay () ;
var j, maille1, an, NumSem, anX, NumSemX,numMoisX,coeff: integer;
    anM,moisM,jourM:word;
    anQ,moisQ,quinzQ,numQuinzQ,quinz,numQuinzQX:word;
    axe1,axe2,axe3,axe4,axe5,axe6,axe7,axe8,axe9,axe10:integer;
    lib1,lib2,lib3,lib4,lib5,lib6,lib7,lib8,lib9,lib10:integer;
    Libelle : hString;
begin
  inherited;
  Libelle := TraduireMemoire('Libellé de ');
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

    if axe1 > 0 then ColCaption[axe1]:=DonneLibelleCAxe(structure,TabNumValAxe[1+1],TabCodeAxe[1+1]);
    if axe2 > 0 then ColCaption[axe2]:=DonneLibelleCAxe(structure,TabNumValAxe[2+1],TabCodeAxe[2+1]);
    if axe3 > 0 then ColCaption[axe3]:=DonneLibelleCAxe(structure,TabNumValAxe[3+1],TabCodeAxe[3+1]);
    if axe4 > 0 then ColCaption[axe4]:=DonneLibelleCAxe(structure,TabNumValAxe[4+1],TabCodeAxe[4+1]);
    if axe5 > 0 then ColCaption[axe5]:=DonneLibelleCAxe(structure,TabNumValAxe[5+1],TabCodeAxe[5+1]);
    if axe6 > 0 then ColCaption[axe6]:=DonneLibelleCAxe(structure,TabNumValAxe[6+1],TabCodeAxe[6+1]);
    if axe7 > 0 then ColCaption[axe7]:=DonneLibelleCAxe(structure,TabNumValAxe[7+1],TabCodeAxe[7+1]);
    if axe8 > 0 then ColCaption[axe8]:=DonneLibelleCAxe(structure,TabNumValAxe[8+1],TabCodeAxe[8+1]);
    if axe9 > 0 then ColCaption[axe9]:=DonneLibelleCAxe(structure,TabNumValAxe[9+1],TabCodeAxe[9+1]);
    if axe10 > 0 then ColCaption[axe10]:=DonneLibelleCAxe(structure,TabNumValAxe[10+1],TabCodeAxe[10+1]);

    if lib1 >= 0 then ColCaption[lib1]:=Libelle+DonneLibelleCAxe(structure,TabNumValAxe[1+1],TabCodeAxe[1+1]);
    if lib2 >= 0 then ColCaption[lib2]:=Libelle+DonneLibelleCAxe(structure,TabNumValAxe[2+1],TabCodeAxe[2+1]);
    if lib3 >= 0 then ColCaption[lib3]:=Libelle+DonneLibelleCAxe(structure,TabNumValAxe[3+1],TabCodeAxe[3+1]);
    if lib4 >= 0 then ColCaption[lib4]:=Libelle+DonneLibelleCAxe(structure,TabNumValAxe[4+1],TabCodeAxe[4+1]);
    if lib5 >= 0 then ColCaption[lib5]:=Libelle+DonneLibelleCAxe(structure,TabNumValAxe[5+1],TabCodeAxe[5+1]);
    if lib6 >= 0 then ColCaption[lib6]:=Libelle+DonneLibelleCAxe(structure,TabNumValAxe[6+1],TabCodeAxe[6+1]);
    if lib7 >= 0 then ColCaption[lib7]:=Libelle+DonneLibelleCAxe(structure,TabNumValAxe[7+1],TabCodeAxe[7+1]);
    if lib8 >= 0 then ColCaption[lib8]:=Libelle+DonneLibelleCAxe(structure,TabNumValAxe[8+1],TabCodeAxe[8+1]);
    if lib9 >= 0 then ColCaption[lib9]:=Libelle+DonneLibelleCAxe(structure,TabNumValAxe[9+1],TabCodeAxe[9+1]);
    if lib10 >= 0 then ColCaption[lib10]:=Libelle+DonneLibelleCAxe(structure,TabNumValAxe[10+1],TabCodeAxe[10+1]);

  end;

  //on affiche : SEM ss/aaaa
  if CocheSemaine then
  begin
    NumSem:= numSemaine(edtDateDeb, an);
    anX := an;

    with TTobViewer(getcontrol('TV')) do
    begin
      maille1 := ColIndex('MAILLE1');
      if maille1 > 0 then
      begin
        for j := maille1 to (maille1 + nbIntervalle - 1)  do
        begin
          numSemX := NumSem - maille1 + j;
          //si le numéro de semaine >52
          //on change d'année
          if numSemX > 52 then
          begin
            numSemX := numSemX - 52;
            if an = anX then anX := an + 1;
          end;

          ColCaption[j] := 'SEM' + inttostr(NumSemX) + '/' + IntToStr(anX);
        end;
      end;
    end;
  end;

  //on affiche : mm_1/aaaa
  if CocheQuinzaine then
  begin
    Date_AnNumQuinz(edtDateDeb,anQ,moisQ,quinzQ,numQuinzQ);
    anX := anQ;
    if quinzQ=1 then quinz:=2
    else quinz:=1;
    with TTobViewer(getcontrol('TV')) do
    begin
      maille1 := ColIndex('MAILLE1');
      if maille1 > 0 then
      begin
        for j := maille1 to (maille1 + nbIntervalle - 1) do
        begin
          numQuinzQX := numQuinzQ - maille1 + j;
          if quinz=1 then
          begin
            quinz:=2;
            numMoisX:=numQuinzQX div 2;
          end
          else
          begin
            quinz:=1;
            numMoisX:=(numQuinzQX div 2)+1;
          end;
          //si le numéro de quinzaine >24
          //on change d'année
          if numQuinzQX > 24 then
          begin
            coeff:=numQuinzQX div 24;
            numQuinzQX := numQuinzQX - (24*coeff);
            if numQuinzQX=1 then quinz:=1;
            if quinz=1 then numMoisX:=(numQuinzQX div 2)+1
            else numMoisX:=numQuinzQX div 2;
            if anQ = anX then anX := anQ + (1*coeff);
          end;

          //ColCaption[j] := MOISI[numMoisX]+'_'+IntToStr(quinz)+'/'+IntToStr(anX);
          if quinz=1 then ColCaption[j] := '01/'+MOISI[numMoisX]+'/'+IntToStr(anX)
          else ColCaption[j] := '16/'+MOISI[numMoisX]+'/'+IntToStr(anX);
        end;
      end;
    end;
  end;


  //on affiche : MOIS mm/aaaa
  if CocheMois then
  begin
    decodeDate(edtDateDeb,anM,moisM,jourM);
    anX := anM;

    with TTobViewer(getcontrol('TV')) do
    begin
      maille1 := ColIndex('MAILLE1');
      if maille1 > 0 then
      begin
        for j := maille1 to (maille1 + nbIntervalle - 1) do
        begin
          numMoisX := moisM - maille1 + j;
          if numMoisX > 12 then
          begin
            coeff:=numMoisX div 12;
            //si mois = decembre  càd (numMoisX mod 12)=0  => mois =12
            //pour trouver l'année on multiplie par un coeff=numMoisX div 12
            if (numMoisX mod 12)=0 then
            begin
              numMoisX:=12;
              anX := anM + (1*coeff)-1;
            end
            else
            begin
              numMoisX := numMoisX - (12*coeff);
              anX := anM + (1*coeff);
            end;
          end;
          ColCaption[j] := TexteMessage[numMoisX] + '-' + IntToStr(anX);
          //ColCaption[j] := '01/'+intToStr(numMoisX)+'/'+IntToStr(anX);
        end;
      end;
    end;
  end;

  if (CocheSemaine=false) and (CocheQuinzaine=false) and (CocheMois=false) and
     (CocheMois445=false) then
  begin
     with TTobViewer(getcontrol('TV')) do
     begin
       maille1 := ColIndex('MAILLE1');
       if maille1 > 0 then ColCaption[maille1]:='Valeur';
     end;
  end;
end ;

procedure TOF_QUFSBPBUDGETR3.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_QUFSBPBUDGETR3.CocheSemaine: boolean;
begin
  result := false;
  if BpInitialise='2' then result := true;
end;

function TOF_QUFSBPBUDGETR3.CocheQuinzaine: boolean;
begin
  result := false;
  if BpInitialise='3' then result := true;
end;

function TOF_QUFSBPBUDGETR3.CocheMois: boolean;
begin
  result := false;
  if BpInitialise='4' then result := true;
end;

function TOF_QUFSBPBUDGETR3.CocheMois445: boolean;
begin
  result := false;
  if BpInitialise='5' then result := true;
end;

procedure TOF_QUFSBPBUDGETR3.killTob;
begin
  if QTBP <> nil then
  begin
    QTBP.Free;
    QTBP:=nil;
  end;
end;

procedure TOF_QUFSBPBUDGETR3.DblClick(Sender: TObject);
var ligne:integer;
    LaDate:TDateTime;
    axe1,axe2,axe3,axe4,axe5,axe6,axe7,axe8,axe9,axe10:integer;
    Tf: Tob;
    i,maille1,j,NbValAff,k:integer;
    maille:double;
    tabMailleH,tabMailleP,tabMaillePR :array [1..12] of double; //CEGID-CCMX ajout tabMaillePR
    TabLibelle:array [0..7] of hString;
    TabAxeS:array [1..10] of hString;
    LibVal:integer;
    LibValS:hstring;
begin
  if (CocheSemaine=false) and (CocheQuinzaine=false) and (CocheMois=false) and
     (CocheMois445=false) then exit;
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
    maille1 := ColIndex('MAILLE1');
    LibVal:= ColIndex('TYPE');
  end;

  if ligne >= 0 then
  begin
    //CEGID-CCMX le 17/11/2006 DEBUT
    if ContextBP<>3 then
    begin
      if (sTypeDA = 'TYPEDA') then QTHPR:=TQTob.create(['LADATE'],['CAHISTO','CAPREVU','CAPRESQUEREAL','CAREALISE'])
      else  //CEGID-CCMX le 17/11/2006 FIN
      QTHPR:=TQTob.create(['LADATE'],['CAHISTO','CAPREVU','CAREALISE']);
    end
    else QTHPR:=TQTob.create(['LADATE'],['CAHISTO','CAPREVU','CAREALISENM','CAREALISE']);

    TabAxeS[1]:='';
    if axe1>=0 then TabAxeS[1]:=TTobViewer(getcontrol('TV')).AsString[axe1,ligne];
    TabAxeS[2]:='';
    if axe2>=0 then TabAxeS[2]:=TTobViewer(getcontrol('TV')).AsString[axe2,ligne];
    TabAxeS[3]:='';
    if axe3>=0 then TabAxeS[3]:=TTobViewer(getcontrol('TV')).AsString[axe3,ligne];
    TabAxeS[4]:='';
    if axe4>=0 then TabAxeS[4]:=TTobViewer(getcontrol('TV')).AsString[axe4,ligne];
    TabAxeS[5]:='';
    if axe5>=0 then TabAxeS[5]:=TTobViewer(getcontrol('TV')).AsString[axe5,ligne];
    TabAxeS[6]:='';
    if axe6>=0 then TabAxeS[6]:=TTobViewer(getcontrol('TV')).AsString[axe6,ligne];
    TabAxeS[7]:='';
    if axe7>=0 then TabAxeS[7]:=TTobViewer(getcontrol('TV')).AsString[axe7,ligne];
    TabAxeS[8]:='';
    if axe8>=0 then TabAxeS[8]:=TTobViewer(getcontrol('TV')).AsString[axe8,ligne];
    TabAxeS[9]:='';
    if axe9>=0 then TabAxeS[9]:=TTobViewer(getcontrol('TV')).AsString[axe9,ligne];
    TabAxeS[10]:='';
    if axe10>=0 then TabAxeS[10]:=TTobViewer(getcontrol('TV')).AsString[axe10,ligne];
    LibValS:='';
    if LibVal>=0 then LibValS:=TTobViewer(getcontrol('TV')).AsString[LibVal,ligne];


    Case ContextBP of
    0,1,2 : begin //Mode-GC-Compta
            //valeur histo
            Tf := QTBP.latob.FindFirst(['TYPE','AXE1','AXE2','AXE3','AXE4','AXE5','AXE6','AXE7','AXE8','AXE9','AXE10'],
                                       ['1-HISTO',TabAxeS[1],TabAxeS[2],TabAxeS[3],TabAxeS[4],TabAxeS[5],TabAxeS[6],TabAxeS[7],TabAxeS[8],TabAxeS[9],TabAxeS[10]],False);
            if Tf <> nil then
            begin
              for i:=1 to 12 do
              begin
                maille:=VALEUR(Tf.GetValue('MAILLE'+(intToStr(i))));
                tabMailleH[i]:=maille;
              end;
            end;

            //valeur Prevu
            Tf := QTBP.latob.FindFirst(['TYPE','AXE1','AXE2','AXE3','AXE4','AXE5','AXE6','AXE7','AXE8','AXE9','AXE10'],
                                       ['2-PREVU',TabAxeS[1],TabAxeS[2],TabAxeS[3],TabAxeS[4],TabAxeS[5],TabAxeS[6],TabAxeS[7],TabAxeS[8],TabAxeS[9],TabAxeS[10]],False);
            if Tf <> nil then
            begin
              for i:=1 to 12 do
              begin
                maille:=VALEUR(Tf.GetValue('MAILLE'+(intToStr(i))));
                tabMailleP[i]:=maille;
              end;
            end;

            //CEGID-CCMX le 17/11/2006 DEBUT
            //valeur Presque Realise
            if (sTypeDA = 'TYPEDA') then
            begin
              Tf := QTBP.latob.FindFirst(['TYPE','AXE1','AXE2','AXE3','AXE4','AXE5','AXE6','AXE7','AXE8','AXE9','AXE10'],
                                         ['4-PRESQUE REALISE',TabAxeS[1],TabAxeS[2],TabAxeS[3],TabAxeS[4],TabAxeS[5],TabAxeS[6],TabAxeS[7],TabAxeS[8],TabAxeS[9],TabAxeS[10]],False);
              if Tf <> nil then
              begin
                for i:=1 to 12 do
                begin
                  maille:=VALEUR(Tf.GetValue('MAILLE'+(intToStr(i))));
                  tabMaillePR[i]:=maille;
                end;
              end;
            end;
            //CEGID-CCMX le 17/11/2006 FIN

            //valeur Realise
            Tf := QTBP.latob.FindFirst(['TYPE','AXE1','AXE2','AXE3','AXE4','AXE5','AXE6','AXE7','AXE8','AXE9','AXE10'],
                                       ['3-REALISE',TabAxeS[1],TabAxeS[2],TabAxeS[3],TabAxeS[4],TabAxeS[5],TabAxeS[6],TabAxeS[7],TabAxeS[8],TabAxeS[9],TabAxeS[10]],False);
            if Tf <> nil then
            begin
              for i:=1 to 12 do
              begin
                maille:=VALEUR(Tf.GetValue('MAILLE'+(intToStr(i))));
                LaDate:=edtDateDeb;
                if CocheMois then LaDate:=PLUSMOIS(edtDateDeb,i-1);
                if CocheSemaine then LaDate:=PlusDate(edtDateDeb,i-1,'S');
                if CocheQuinzaine then LaDate:=StrToDateTime(TTobViewer(getcontrol('TV')).ColCaption[maille1+i-1]);
                //CEGID-CCMX le 17/11/2006 ajout tabMaillePR[i]   //,tabMaillePR[i]
                if (sTypeDA ='TYPEDA') then QTHPR.addValeur([LaDate],[tabMailleH[i],tabMailleP[i],tabMaillePR[i],maille])
                else QTHPR.addValeur([LaDate],[tabMailleH[i],tabMailleP[i],maille]);
              end;
            end;
            titreGraph:='Axes : '+TabAxeS[1];
            for i:= 2 to 10 do
            begin
              if TabAxeS[i]<>'' then titreGraph:=titreGraph+'/'+TabAxeS[i];
            end;
          end;
      3 : begin
            LibValAff(edtSession,TabLibelle);
            NbValAff := StrToInt(TabLibelle[0]);
            k:=1;
            for j := 1 to NbValAff do
            begin
              if pos(TabLibelle[j],LibValS) > 0 then
              begin
                k:=j;
                break;
              end;
            end;

            //valeur Théorique
            Tf := QTBP.latob.FindFirst(['TYPE','AXE1','AXE2','AXE3','AXE4','AXE5','AXE6','AXE7','AXE8','AXE9','AXE10'],
                                       [TabLibelle[k]+TraduireMemoire(' [Théorique]'),TabAxeS[1],TabAxeS[2],TabAxeS[3],TabAxeS[4],TabAxeS[5],TabAxeS[6],TabAxeS[7],TabAxeS[8],TabAxeS[9],TabAxeS[10]],False);
            if Tf <> nil then
            begin
              for i:=1 to 12 do
              begin
                maille:=VALEUR(Tf.GetValue('MAILLE'+(intToStr(i))));
                tabMailleH[i]:=maille;
              end;
            end;

            //valeur Budget N
            Tf := QTBP.latob.FindFirst(['TYPE','AXE1','AXE2','AXE3','AXE4','AXE5','AXE6','AXE7','AXE8','AXE9','AXE10'],
                                       [TabLibelle[k]+TraduireMemoire(' [Budget N]'),TabAxeS[1],TabAxeS[2],TabAxeS[3],TabAxeS[4],TabAxeS[5],TabAxeS[6],TabAxeS[7],TabAxeS[8],TabAxeS[9],TabAxeS[10]],False);
            if Tf <> nil then
            begin
              for i:=1 to 12 do
              begin
                maille:=VALEUR(Tf.GetValue('MAILLE'+(intToStr(i))));
                tabMailleP[i]:=maille;
              end;
            end;

            //valeur Réalisé N-1
            if Tf <> nil then
            begin
              Tf := QTBP.latob.FindFirst(['TYPE','AXE1','AXE2','AXE3','AXE4','AXE5','AXE6','AXE7','AXE8','AXE9','AXE10'],
                                         [TabLibelle[k]+TraduireMemoire(' [Réalisé N-1]'),TabAxeS[1],TabAxeS[2],TabAxeS[3],TabAxeS[4],TabAxeS[5],TabAxeS[6],TabAxeS[7],TabAxeS[8],TabAxeS[9],TabAxeS[10]],False);
              if Tf <> nil then
              begin
                for i:=1 to 12 do
                begin
                  maille:=VALEUR(Tf.GetValue('MAILLE'+(intToStr(i))));
                  tabMaillePR[i]:=maille;
                end;
              end;
            end;

            //valeur Realise N
            Tf := QTBP.latob.FindFirst(['TYPE','AXE1','AXE2','AXE3','AXE4','AXE5','AXE6','AXE7','AXE8','AXE9','AXE10'],
                                      [TabLibelle[k]+TraduireMemoire(' [Réalisé N]'),TabAxeS[1],TabAxeS[2],TabAxeS[3],TabAxeS[4],TabAxeS[5],TabAxeS[6],TabAxeS[7],TabAxeS[8],TabAxeS[9],TabAxeS[10]],False);
            if Tf <> nil then
            begin
              for i:=1 to 12 do
              begin
                maille:=VALEUR(Tf.GetValue('MAILLE'+(intToStr(i))));
                LaDate:=edtDateDeb;
                if CocheMois then LaDate:=PLUSMOIS(edtDateDeb,i-1);
                if CocheSemaine then LaDate:=PlusDate(edtDateDeb,i-1,'S');
                if CocheQuinzaine then LaDate:=StrToDateTime(TTobViewer(getcontrol('TV')).ColCaption[maille1+i-1]);
                QTHPR.addValeur([LaDate],[tabMailleH[i],tabMailleP[i],tabMaillePR[i],maille]);
              end;
            end;
            titreGraph:='Axes : '+TabAxeS[1];
            for i:= 2 to 10 do
            begin
              if TabAxeS[i]<>'' then titreGraph:=titreGraph+'/'+TabAxeS[i];
            end;
            titreGraph:=titreGraph+' ['+TabLibelle[k]+']';
          end;

    end; //CASE

    AglLanceFiche('Q', 'QUFGBPGRAPHHPR', '', '', '');
    QTHPR.free;
  end;
end;

{ EVI / Récupération des critères - multiselection}
procedure TOF_QUFSBPBUDGETR3.DonneFiltre;
var i,Dot,DotTemp:integer;
AxeTemp:hString;
edtAxe1,edtAxe2,edtAxe3,edtAxe4,edtAxe5:hString;
edtAxe6,edtAxe7,edtAxe8,edtAxe9,edtAxe10:hString;
begin
  codeFiltre:='';

  edtAxe1:=GetControlText('EDTAXE1');
  edtAxe2:=GetControlText('EDTAXE2');
  edtAxe3:=GetControlText('EDTAXE3');
  edtAxe4:=GetControlText('EDTAXE4');
  edtAxe5:=GetControlText('EDTAXE5');
  edtAxe6:=GetControlText('EDTAXE6');
  edtAxe7:=GetControlText('EDTAXE7');
  edtAxe8:=GetControlText('EDTAXE8');
  edtAxe9:=GetControlText('EDTAXE9');
  edtAxe10:=GetControlText('EDTAXE10');

  Dot :=0;

  //-----------------> ORLI
  if BPOkOrli
  then
  begin
    if edtAxe1<>'' then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[2]+' like "'+edtAxe1+'%" ';
    if edtAxe2<>'' then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[3]+' like "'+edtAxe2+'%" ';
    if edtAxe3<>'' then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[4]+' like "'+edtAxe3+'%" ';
    if edtAxe4<>'' then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[5]+' like "'+edtAxe4+'%" ';
    if edtAxe5<>'' then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[6]+' like "'+edtAxe5+'%" ';
    if edtAxe6<>'' then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[7]+' like "'+edtAxe6+'%" ';
    if edtAxe7<>'' then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[8]+' like "'+edtAxe7+'%" ';
    if edtAxe8<>'' then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[9]+' like "'+edtAxe8+'%" ';
    if edtAxe9<>'' then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[10]+' like "'+edtAxe9+'%" ';
    if edtAxe10<>'' then codeFiltre:=codeFiltre+' AND QBQ_VALAXECT'+TabNumValAxe[11]+' like "'+edtAxe10+'%" ';
  end
  //<----------------- ORLI
  else
  begin
    { EVI / Récupération des critères dans le cas multiselection }
    codefiltre:=' AND QBQ_CODESESSION="'+edtSession+'" ';
    for i:=1 to 10 do
    begin
      if (GetControlText('EDTAXE'+intToStr(i))<>'') and (Pos(';',GetControlText('EDTAXE'+intToStr(i)))<>0)
      then
      begin
        Dot := VerifPos(true,GetControlText('EDTAXE'+intToStr(i)), ';',Dot);
        AxeTemp := Copy(GetControlText('EDTAXE'+intToStr(i)),1,Dot-2);
        codefiltre:= codefiltre+' AND QBQ_VALAXECT'+intToStr(i)+' in ("'+AxeTemp+'"';
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
      if Uppercase(GetControlText('EDTAXE'+intToStr(i)))<>Uppercase(TraduireMemoire('<<Tous>>')) then codefiltre:=codefiltre+' AND QBQ_VALAXECT'+intToStr(i)+' like "'+GetControlText('EDTAXE'+intToStr(i))+'%" ';
    end;
  end;
if BPOkOrli then codeFiltre:=codeFiltre+' AND QBQ_CODESESSION="'+edtSession+'" ';
end;

{ EVI / Trap de la touche ESC : éviter de sortir pendant la génération }
procedure TOF_QUFSBPBUDGETR3.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 if (Key = VK_ESCAPE) and (InProgress=true) then Key := 0;
 if( (Key = VK_F9) and Assigned(BCherche) ) then BCherche.Click;
end;

Initialization
  registerclasses ( [ TOF_QUFSBPBUDGETR3 ] ) ;
end.
