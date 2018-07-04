{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 21/07/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFVBPCALENDJOURM ()
Mots clefs ... : TOF;QUFVBPCALENDJOURM
*****************************************************************}
Unit QUFVBPCALENDJOURM_TOF ;

Interface

uses UTOF,uTob,HEnt1,Graphics,Grids,BPMaille,HCtrls;

Type
  TOF_QUFVBPCALENDJOURM = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
private
     TabVerifPrct,TabSumPeriodPrevu:array of double;
     TabNumPeriod:array of hString;
     TabAxe:array[1..10] of hString;
     DateDebSemR,DateDebSession,DateFinSession,DateFinPeriod,DateDebPeriod,DateDebSem,DateDebRef,DateFinRef:TDateTime;
     LeapFeb,Decalage,BPInitialise,LigneEnter,NumPeriod,numSem,nbJour,nbJourPeriod,nbSem:integer;
     ModeFiche,CodeSession,Calendrier,ValAff,ValCell,ValidationAuto:hString;
     SommeHistoArbre,SommePrevuArbre:double;
     IncPeriod,CellValid,OkValid,StateFiche:boolean;
     laTob,TobUpdate:Tob;
     Mi:TMaille;
     GridCalend:THGrid;
     RadioGrpMode:THRadioGroup;
     procedure GetCellCanvas( ACol, ARow : Longint; Canvas : TCanvas; AState : TGridDrawState );
     procedure RemplitGrid;
     procedure MAJTableGrid;
     procedure Resize( Sender : TObject );
     procedure BtnSolde(Sender: TObject);
     procedure BImprimerOnClick(Sender: TObject);
     procedure CellEnter(Sender: TObject; var ACol,ARow: Longint;var Cancel: Boolean);
     procedure CellExit(Sender: TObject; var ACol,ARow: Longint;var Cancel: Boolean);
     procedure RowEnter( Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean );
     procedure RecalculTot(Verif,Prct:Boolean;ligne:integer);
     procedure RecalculTotPeriode;
     procedure CellEdit(Sender: TObject; ACol, ARow: Integer; const Value: hString);
     procedure DonneDetailPeriod(Jour:TDateTime;i:integer);
     procedure CellClick(Sender: TObject);
     procedure PopupArbre(Sender: TObject);
     procedure RecalcReliq_OnClick(Sender: TObject);     
     procedure RadioGrpModeState(Sender: TObject);
  end ;

Implementation

Uses Classes,sysutils,HTB97,HmsgBox,ComCtrls,
     Windows,messages,Forms,Menus,Math,
     {$IFNDEF EAGLCLIENT}EdtREtat,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     {$ELSE}UtileAGL,
     {$ENDIF}
     BPBasic,UUtil,BPFctSession,BPUtil;

function NumMois( Date : TDateTime ) : integer;
var Annee, Mois, Jour : Word;
begin
  DecodeDate( Date, Annee, Mois, Jour );
  Result := Mois;
end;

procedure TOF_QUFVBPCALENDJOURM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPCALENDJOURM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPCALENDJOURM.OnUpdate ;
var i:integer;
begin
  Inherited ;
  MAJTableGrid;
  StateFiche := false;


  if OkValid
  then
  begin
    if ModeFiche = 'CALENDCONSULT'
    then for i:=0 to TobUpdate.detail.count-1 do TobUpdate.Detail[i].UpdateDB(true)
    else for i:=0 to TobUpdate.detail.count-1 do TobUpdate.Detail[i].InsertDB(nil,true);
  end;
end ;

procedure TOF_QUFVBPCALENDJOURM.OnLoad ;
begin
  Inherited ;
  RemplitGrid;
  RecalculTot(true,true,0);
  if ValidationAuto='TRUE' then
  begin
    Update ;
    PostMessage(Ecran.Handle, WM_CLOSE, 0, 0);
  end
  else
  begin
    Ecran.WindowState := wsNormal ;
    Resize(self);
  end;
end ;

procedure TOF_QUFVBPCALENDJOURM.DonneDetailPeriod(Jour:TDateTime;i:integer);
var AnSemDeb,numPer,anPer:integer;
    anQ,moisQ,QuinzQ,NumQuinzQ:word;
    DateDebSemTemp:TDateTime;
begin
  case BPInitialise of
    2 : //Semaine
      begin
        if i <> 0 then
        begin
          DateFinPeriod:=jour;
          numPeriod:=Ceil(i/7);
          if Ceil(i/7)<>Ceil((i+1)/7) then
          begin
            DateFinPeriod:=jour;
            DateDebPeriod:=jour;
          end
          else
          begin
            DateDebPeriod:=jour+1;
            DateFinPeriod:=jour+1;
          end;
        end;
        numSem:=NumSemaine(jour,AnSemDeb);
        DateDebSemTemp:=PremierJourSemaine(numSem,AnSemDeb);
        nbJour := round(jour-DateDebSemTemp);
        nbJourPeriod:=6;
      end;
    3 : //Quinzaine
      begin
        Date_AnNumQuinz(jour,anQ,moisQ,QuinzQ,NumQuinzQ);
        if QuinzQ=1 then
        begin
          DateDebPeriod:=DebutDeMois(jour);
          DateFinPeriod:=EncodeDate(AnQ,moisQ,15)
        end
        else
        begin
          DateDebPeriod:=EncodeDate(AnQ,moisQ,16);
          DateFinPeriod:=FinDeMois(jour);
        end;
        IncPeriod:=false;
        nbJour := round(jour-DateDebPeriod);
        nbJourPeriod:=round(DateFinPeriod-DateDebPeriod);
      end;
    4 : //Mois
      begin
        DateDebPeriod:= DebutDeMois(jour);
        DateFinPeriod:= FinDeMois(jour);
        IncPeriod:=false;
        nbJour := round(jour-DateDebPeriod);
        nbJourPeriod:=round(DateFinPeriod-DateDebPeriod) ;
      end;
    6 : //Trimestre
      begin
        DonneNumTrimestre(jour,numPer,anPer);
        DonneDateDebFinTrimestre(numPer,anPer,DateDebPeriod,DateFinPeriod);
        IncPeriod:=false;
        nbJour := round(jour-DateDebPeriod);
        nbJourPeriod:=round(DateFinPeriod-DateDebPeriod);
      end;
    7 : //Quadrimestre
      begin
        DonneNumQuadrimestre(jour,numPer,anPer);
        DonneDateDebFinQuadrimestre(numPer,anPer,DateDebPeriod,DateFinPeriod);
        IncPeriod:=false;
        nbJour := round(jour-DateDebPeriod);
        nbJourPeriod:=round(DateFinPeriod-DateDebPeriod);
      end;
    1,5 :  { EVI / ToDo : Mois 4-4-5 }
      begin
        // SetControlText('TH_PERIODE','') ;
        LigneEnter:=1;
        nbJour:=0;
        nbJourPeriod:=GridCalend.RowCount-2;
      end;
    //else //CASE
  end; // CASE
end;



procedure TOF_QUFVBPCALENDJOURM.RemplitGrid;
var i,j,imax,DebPeriod,k,numPer,anPer,nivMax,numPerDeb,anPerDeb:integer;
    TabSumPeriodHisto,TabSumPeriodHistoDecale :array of double;
    SommePrctHistorique,SommePrctPrevu,MaxTab,MaxTabTemp:double;
    SommeVerifPrctPeriod,SommePeriodTypePrevu: double;
    DateSQL,jour:TDateTime;
    anQ,moisQ,QuinzQ,NumQuinzQ:word;
    days: array[1..7] of hString;
    ValeurAxe,SQLValAff,AxeCube:hString;
    Init:boolean;
    Q:TQuery;
begin
  SetLength(TabSumPeriodHisto,nbSem+1);
  SetLength(TabSumPeriodHistoDecale,nbSem+1);

  days[1] := TraduireMemoire('Di');
  days[2] := TraduireMemoire('Lu');
  days[3] := TraduireMemoire('Ma');
  days[4] := TraduireMemoire('Me');
  days[5] := TraduireMemoire('Je');
  days[6] := TraduireMemoire('Ve');
  days[7] := TraduireMemoire('Sa');

  GridCalend.RowCount:=2;

  SommePrctHistorique:=0;
  SommeVerifPrctPeriod:=0;
  SommePeriodTypePrevu:=0;
  DebPeriod:=1;

  for i:=1 to nbSem*7 do
  begin
    jour:=PLUSDATE(DateDebSemR,i-1,'J');

    case BPInitialise of
      2 : //Semaine
        begin
          numSem:=NumSemaine(jour);
          Latob.Detail[i-1].putValue('JOUR',format(TraduireMemoire('Sem  ')+'%s : %-2s %s ',[MetZero(IntToStr(numSem),2),days[DayOfWeek(jour)],DateTimeToStr(jour)]));
        end;

      3 : //Quinzaine
        begin
          Date_AnNumQuinz(jour,anQ,moisQ,QuinzQ,NumQuinzQ);
          Latob.Detail[i-1].putValue('JOUR',format(TraduireMemoire('Quin ')+'%s : %-2s %s ',[MetZero(IntToStr(NumQuinzQ),2),days[DayOfWeek(jour)],DateTimeToStr(jour)]));
        end;

      4 : //Mois
        begin
          numSem:=NumMois(jour);
          Latob.Detail[i-1].putValue('JOUR',format(TraduireMemoire('Mois ')+'%s : %-2s %s ',[MetZero(IntToStr(numSem),2),days[DayOfWeek(jour)],DateTimeToStr(jour)]));
        end;
      6 : //Trimestre
        begin
          DonneNumTrimestre(jour,numPer,anPer);
          Latob.Detail[i-1].putValue('JOUR',format(TraduireMemoire('Trim ')+'%s : %-2s %s ',[MetZero(IntToStr(numPer),2),days[DayOfWeek(jour)],DateTimeToStr(jour)]));
        end;
      7 : //Quadrimestre
        begin
          DonneNumQuadrimestre(jour,numPer,anPer);
          Latob.Detail[i-1].putValue('JOUR',format(TraduireMemoire('Quad ')+'%s : %-2s %s ',[MetZero(IntToStr(numPer),2),days[DayOfWeek(jour)],DateTimeToStr(jour)]));
        end;
      else //CASE { EVI / ToDo : Mois 4-4-5 }
        begin
          numSem:=NumSemaine(jour);
          Latob.Detail[i-1].putValue('JOUR',format(TraduireMemoire('Sem  ')+'%s : %-2s %s ',[MetZero(IntToStr(numSem),2),days[DayOfWeek(jour)],DateTimeToStr(jour)]));
        end;
    end;  //CASE
  end;

  { EVI / Remplissage de % initiaux et % saisies -> JOUR }
  if (SommeHistoArbre<>0)
  then
  begin
    MaxTab:=0;imax:=1;
    for i:=1 to nbSem*7 do
    begin
      jour:=PLUSDATE(DateDebSemR,i-1,'J');
      MaxTabTemp := (Latob.Detail[i-1].getValue('HISTO')/SommeHistoArbre)*100;
      Latob.Detail[i-1].putValue('PRCTGLOBAL',MaxTabTemp);
      if (DateDebSession<=(jour+decalage))  AND  ((jour+decalage)<=DateFinSession)
      then SommePrctHistorique:=SommePrctHistorique+Latob.Detail[i-1].getValue('PRCTGLOBAL');

      if MaxTabTemp>MaxTab then
      begin
        MaxTab:=MaxTabTemp;
        imax:=i;
      end;
    end;

    { EVI / Rétablit pourcentage à 100% }
    if (Arrondi(SommePrctHistorique,4)<>100.0000) AND (Arrondi(SommePrctHistorique,4)<>0.0000) then
    Latob.Detail[imax-1].putValue('PRCTGLOBAL',Latob.Detail[imax-1].getValue('PRCTGLOBAL')+(100-SommePrctHistorique));

    { EVI / Si il ne s'agit pas d'une initialisation par calendrier type
      Ou si on n'est pas en consultation }
    if (ModeFiche<>'CALENDTYPE') AND (ModeFiche<>'CALENDCONSULT') then
    begin
      for i:=1 to nbSem*7 do
      Latob.Detail[i-1].putValue('PRCTPREVU',Latob.Detail[i-1].getValue('PRCTGLOBAL'));
    end;
  end
  else
//  HShowmessage('1;Attention;Aucune donnée historique n''existe pour la sélection.'+#13#10+' La colonne des pourcentages initiaux sera vide.;E;O;', '', '');
  begin
    MaxTab:=0;imax:=1;
    for i:=1 to nbSem*7 do
    begin
      jour:=PLUSDATE(DateDebSemR,i-1,'J');
      MaxTabTemp := (Latob.Detail[i-1].getValue('PREVU')/SommePrevuArbre)*100;
      Latob.Detail[i-1].putValue('PRCTGLOBAL',MaxTabTemp);
      if (DateDebSession<=(jour+decalage))  AND  ((jour+decalage)<=DateFinSession)
      then SommePrctHistorique:=SommePrctHistorique+Latob.Detail[i-1].getValue('PRCTGLOBAL');

      if MaxTabTemp>MaxTab then
      begin
        MaxTab:=MaxTabTemp;
        imax:=i;
      end;
    end;

    { EVI / Rétablit pourcentage à 100% }
    if (Arrondi(SommePrctHistorique,4)<>100.0000) AND (Arrondi(SommePrctHistorique,4)<>0.0000) then
    Latob.Detail[imax-1].putValue('PRCTGLOBAL',Latob.Detail[imax-1].getValue('PRCTGLOBAL')+(100-SommePrctHistorique));

    { EVI / Si il ne s'agit pas d'une initialisation par calendrier type
      Ou si on n'est pas en consultation }
    if (ModeFiche<>'CALENDTYPE') AND (ModeFiche<>'CALENDCONSULT') then
    begin
      for i:=1 to nbSem*7 do
      Latob.Detail[i-1].putValue('PRCTPREVU',Latob.Detail[i-1].getValue('PRCTGLOBAL'));
    end;
  end;

  { EVI / Récupération de la répartition par période Historique et Prevu dans l'arbre }
  Case ContextBP of
    0,1 : begin //Mode-GC
          if ValAff = 'PTC' then SQLValAff:='SUM(QBR_REF1),SUM(QBR_OP1),QBR_VALEURAXE';              //CA TTC net
          if ValAff = 'QTE' then SQLValAff:='SUM(QBR_QTEREF),SUM(QBR_QTEC),QBR_VALEURAXE';           //Quantité
          if ValAff = 'PHT' then SQLValAff:='SUM(QBR_REF2),SUM(QBR_OP2),QBR_VALEURAXE';              //CA HT net
          if ValAff = 'UHT' then SQLValAff:='SUM(QBR_REF3),SUM(QBR_OP3),QBR_VALEURAXE';              //CA HT avant remise
          if ValAff = 'UTC' then SQLValAff:='SUM(QBR_REF4),SUM(QBR_OP4),QBR_VALEURAXE';              //CA TTC avant remise
          if ValAff = 'PAH' then SQLValAff:='SUM(QBR_REF5),SUM(QBR_OP5),QBR_VALEURAXE';              //Cumul achat PMAP
          if ValAff = 'MAR' then SQLValAff:='SUM(QBR_REF6),SUM(QBR_OP6),QBR_VALEURAXE';              //Marge
        end;
    2 : begin //Compta
          if ValAff = 'DC1' then SQLValAff:='SUM( QBR_REF1),SUM(QBR_OP1),QBR_VALEURAXE';
          if ValAff = 'CD1' then SQLValAff:='SUM(QBR_REF2),SUM(QBR_OP2),QBR_VALEURAXE';
        end;
  end; //CASE

  i:=1;
  AxeCube := '';
  While TabAxe[i] <> '' do
  begin
    if TabAxe[i]='[---]' then AxeCube := AxeCube + ' AND (QBR_VALAXENIV'+IntToStr(i)+' is null OR QBR_VALAXENIV'+IntToStr(i)+'="")'
    else AxeCube := AxeCube + ' AND QBR_VALAXENIV'+IntToStr(i)+'="'+TabAxe[i]+'"';
    i:= i+1;
  end;

  nivMax:=ChercheNivMax(codeSession)+1;

  Q:=MOpenSql('SELECT '+SQLValAff+' FROM QBPARBRE WHERE QBR_CODESESSION="'+CodeSession+'"'+AxeCube+
              ' AND QBR_NIVEAU="'+IntToStr(nivMax)+'" GROUP BY QBR_VALEURAXE ORDER BY QBR_VALEURAXE','QUFVBPCALENDJOURM_TOF (RemplitGrid).',true);

  While not Q.eof
  do
  begin
    ValeurAxe:=Q.Fields[2].AsString;
    DateSQL:=StrToDateTime(ValeurAxe);

    case BPInitialise of
          2 : NumPeriod:=NbSemaineIntervalle(DateDebSession,DateSQL);
          3 : NumPeriod:=NbQuinzaineIntervalle(DateDebSession,DateSQL);
          4 : NumPeriod:=NbMoisIntervalle(DateDebSession,DateSQL);
          6 : //Trimestre
            begin
              DonneNumTrimestre(DateDebSession,numPerDeb,anPerDeb);
              DonneNumTrimestre(DateSQL,numPer,anPer);
              NumPeriod:=numPer-NumPerDeb+((anPer-anPerDeb)*4)+1;
            end;
          7 : //Quadrimestre
            begin
              DonneNumQuadrimestre(DateDebSession,numPerDeb,anPerDeb);
              DonneNumQuadrimestre(DateSQL,numPer,anPer);
              NumPeriod:=numPer-NumPerDeb+((anPer-anPerDeb)*3)+1;
            end;
          //else //CASE { EVI / ToDo : Mois 4-4-5 }
        end; // CASE

    TabSumPeriodHisto[NumPeriod]:=Q.Fields[0].asFloat;
    TabSumPeriodPrevu[NumPeriod]:=Q.Fields[1].asFloat;
    Q.Next;
  end;
  ferme(Q);

  { EVI / Recalcul des pourcentages initiaux }
  Init:=false;
  IncPeriod:=true;
  numPeriod:=1;

  //Calcul des sommes décalées pour chaque période
  for i:=1 to nbSem*7 do
  begin
    jour:=PLUSDATE(DateDebSemR,i-1,'J');
    if (DateDebSession<=jour)  AND  (jour<=DateFinSession) then
    begin
      if IncPeriod=true then DonneDetailPeriod(jour,i);
      if not (jour=DateFinPeriod) then TabSumPeriodHistoDecale[numPeriod]:=TabSumPeriodHistoDecale[numPeriod]+latob.detail[i-1].getValue('HISTO')
      else
      begin
        TabSumPeriodHistoDecale[numPeriod]:=TabSumPeriodHistoDecale[numPeriod]+latob.detail[i-1].getValue('HISTO');
        numPeriod:=numPeriod+1;
        IncPeriod:=true;
      end;
    end;
  end;

  IncPeriod:=true;
  numPeriod:=1;

  for i:=1 to nbSem*7 do
  begin
    Latob.Detail[i-1].putValue('PRCTHISTO',Latob.Detail[i-1].getValue('PRCTGLOBAL'));
    jour:=PLUSDATE(DateDebSemR,i-1,'J');
    if (DateDebSession<=jour)  AND  (jour<=DateFinSession) then
    begin

      if Init=false then DebPeriod:=i;
      Init:=true;

      if IncPeriod=true then DonneDetailPeriod(jour,i);

      if not (jour=DateFinPeriod) then
      begin
        if TabSumPeriodHistoDecale[numPeriod]<>0 then
        Latob.Detail[i-1].putValue('PRCTHISTO',Arrondi(((Latob.Detail[i-1].getValue('HISTO')/TabSumPeriodHistoDecale[numPeriod])*100),4))
        else Latob.Detail[i-1].putValue('PRCTHISTO',0);
        SommeVerifPrctPeriod:=SommeVerifPrctPeriod+ Latob.Detail[i-1].getValue('PRCTHISTO');
        If ModeFiche = 'CALENDTYPE' then SommePeriodTypePrevu:=SommePeriodTypePrevu+Latob.Detail[i-1].getValue('PRCTPREVU');
      end
      else
      begin
        if TabSumPeriodHistoDecale[numPeriod]<>0 then
        Latob.Detail[i-1].putValue('PRCTHISTO',Arrondi(((Latob.Detail[i-1].getValue('HISTO')/TabSumPeriodHistoDecale[numPeriod])*100),4))
        else Latob.Detail[i-1].putValue('PRCTHISTO',0);
        SommeVerifPrctPeriod:=SommeVerifPrctPeriod+ Latob.Detail[i-1].getValue('PRCTHISTO');

        { EVI / Verif % = 100 }
        MaxTab:=0;
        imax:=DebPeriod;
        if (Arrondi(SommeVerifPrctPeriod,4)<>100.0000) AND (Arrondi(SommeVerifPrctPeriod,4)<>0.0000) then
        begin
          for j:= DebPeriod to i do
          begin
            MaxTabTemp := Latob.Detail[j-1].getValue('PRCTHISTO');
            if MaxTabTemp>MaxTab then
            begin
              MaxTab:=MaxTabTemp;
              imax:=j;
            end;
          end;
          Latob.Detail[imax-1].putValue('PRCTHISTO',Latob.Detail[imax-1].getValue('PRCTHISTO')+(100-SommeVerifPrctPeriod))
        end;

        SommeVerifPrctPeriod:=0;

        If ModeFiche = 'CALENDTYPE' then
        begin
          SommePeriodTypePrevu:=SommePeriodTypePrevu+latob.detail[i-1].getValue('PRCTPREVU');
          for k:= DebPeriod to i do
          begin
            Latob.Detail[k-1].putValue('PRCTPREVU',Arrondi(Latob.Detail[k-1].getValue('PRCTPREVU')/SommePeriodTypePrevu*100,4));
            SommeVerifPrctPeriod:=SommeVerifPrctPeriod+ Latob.Detail[k-1].getValue('PRCTPREVU');
          end;

          { EVI / Verif % = 100 }
          MaxTab:=0;
          imax:=DebPeriod;
          if (Arrondi(SommeVerifPrctPeriod,4)<>100.0000) AND (Arrondi(SommeVerifPrctPeriod,4)<>0.0000) then
          begin
            for j:= DebPeriod to i do
            begin
              MaxTabTemp := Latob.Detail[j-1].getValue('PRCTPREVU');
              if MaxTabTemp>MaxTab then
              begin
                MaxTab:=MaxTabTemp;
                imax:=j;
              end;
            end;
            Latob.Detail[imax-1].putValue('PRCTPREVU',Latob.Detail[imax-1].getValue('PRCTPREVU')+(100-SommeVerifPrctPeriod))
          end;
        end;

        DebPeriod:=i+1;
        SommePeriodTypePrevu:=0;
        SommeVerifPrctPeriod:=0;
        IncPeriod:=true;
        numPeriod:=numPeriod+1;
      end;
    end
    else
    begin
      Latob.Detail[i-1].putValue('PRCTHISTO',0)
    end;
  end;

  SommePrctHistorique:=0;
  SommePrctPrevu:=0;
  for i:=1 to nbSem*7 do
  begin
    if (ModeFiche<>'CALENDTYPE') AND (ModeFiche<>'CALENDCONSULT') then
    Latob.Detail[i-1].putValue('PRCTPREVU',Latob.Detail[i-1].getValue('PRCTHISTO'));

    jour:=PLUSDATE(DateDebSemR,i-1,'J');
    if (DateDebSession<=jour)  AND  (jour<=DateFinSession)
    then
    begin
      SommePrctHistorique:=SommePrctHistorique+Latob.Detail[i-1].getValue('PRCTHISTO');
      SommePrctPrevu:=SommePrctPrevu+Latob.Detail[i-1].getValue('PRCTPREVU');
    end else Latob.Detail[i-1].putValue('PRCTPREVU','0,0000');
  end;

  SetControlText('EDTTOTAL',Format('%10.2f',[SommePrctHistorique]));     //clYellow
  THNumEdit(getcontrol('EDTHISTO')).Value:=SommeHistoArbre;
  SetControlText('EDTPREVORIG',Format('%10.2f',[SommePrevuArbre]));     //clYellow
  laTob.PutGridDetail(GridCalend,false,false,'JOUR;PRCTHISTO;PRCTPREVU;COM;HISTO;PREVU');
end;

procedure TOF_QUFVBPCALENDJOURM.MAJTableGrid;
var jour,ListFalsePeriode:hString;
    jourSem:TDateTime;
    i,j:integer;
    Tf:Tob;
begin
  OkValid:=true;
  ListFalsePeriode:='';
  { EVI / Pour chaque période controle du total du prct }
  for j:=1 to nbSem do
  begin
    if (Arrondi(TabVerifPrct[j],4)<>100.0000) AND (Arrondi(TabVerifPrct[j],4)<>0.0000) then
    begin
      OKValid:=false;
      laTob.GetLigneGrid(GridCalend,(j*7),';JOUR');
      jour:=laTob.getValue('JOUR');
      ListFalsePeriode := ListFalsePeriode + ' ' + TabNumPeriod[j] + ' - Somme = '+ FloatToStr(Arrondi(TabVerifPrct[j],4)) + #13#10;
    end;
  end;

  if OkValid=false then
  begin
    LastErrorMsg := TraduireMemoire('La somme des pourcentages doit être égale à 100#13#10 '+
                                    '(ou 0 pour une période sans vente)#13#10')+ ListFalsePeriode;
    if Ecran.WindowState = wsMinimized then Ecran.WindowState := wsNormal ;
  end
  else
  begin

    for j:=1 to nbSem do
    begin
      jourSem:=0;
      for i:=0 to 6 do
      begin
        laTob.GetLigneGrid(GridCalend,i+((j-1)*7)+1,';JOUR');
        jour:=laTob.getValue('JOUR');
        if i=0 then jourSem:=StrToDatetime(trim(copy(jour,13,11)));
      end;

      Tf := Tob.Create('QBPDETCALENDREP', TobUpdate, -1);
      Tf.putvalue('QBE_CALENDREP', CodeSession);
      Tf.putvalue('QBE_DATEDEBSEM',jourSem);
      for i := 0 to 6 do
      begin
        if Latob.Detail[((j-1)*7)+i].getString('PRCTPREVU') = '' then
        begin
          Tf.putvalue('QBE_PRCTJOUR'+IntToStr(i+1),0)
        end
        else Tf.putvalue('QBE_PRCTJOUR'+IntToStr(i+1),Latob.Detail[((j-1)*7)+i].getValue('PRCTPREVU'));
        Tf.putvalue('QBE_COMMENTAIRE'+IntToStr(i+1),Latob.Detail[((j-1)*7)+i].getValue('COM'));
        if Latob.Detail[((j-1)*7)+i].getString('PRCTGLOBAL') = '' then
        begin
          Tf.putvalue('QBE_PRCTHISTO'+IntToStr(i+1),0)
        end
        else Tf.putvalue('QBE_PRCTHISTO'+IntToStr(i+1),Latob.Detail[((j-1)*7)+i].getValue('PRCTGLOBAL'));
      end;
      for i := 1 to 10 do Tf.putvalue('QBE_VALAXENIV'+IntToStr(i),TabAxe[i]);

    end;
  end;
end;

procedure TOF_QUFVBPCALENDJOURM.OnArgument (S : String ) ;
var i,j,AnSemDeb,AnSemFin,NumSemDeb,NumSemFin:integer;
    SQLValAff,Title,AxeCube:hString;
    DateFinSem,DateCt,DateCourante,DateDebC,DateFinC,DatePrevu:TDateTime;
    ListMailleJour:TListMaille;
    SommeHisto,SommePrevu:double;
    Q:TQuery;
    Tf:Tob;
    DecalageR,DecalageC:integer;
begin
  Inherited ;
  DateDebC:=0;
  DateFinC:=0;
  DateDebRef:=0;
  DateFinRef:=0;
  OkValid:=true;
  LeapFeb := 0;
  GridCalend:=THGrid(getcontrol('GRIDLOI'));
  RadioGrpMode:=THRadioGroup(GetControl('RadioGrpMode'));
  RadioGrpMode.OnClick := RadioGrpModeState;
  CodeSession:=TrouveArgument(S,'SESSION','');
  Calendrier:=Trouveargument(S,'CALEND','');
  ModeFiche:=Trouveargument(S,'MODEFICHE','');
  BPInitialise:=StrToInt(SessionBPInitialise(CodeSession));
  SommeHistoArbre:=Trouveargument(S,'HISTO','');
  SommePrevuArbre:=Trouveargument(S,'PREVU','');
  ValidationAuto:=TrouveArgument(S,'VALIDATIONAUTO','');
  if ModeFiche='CALENDCONSULT' then StateFiche := false else StateFiche := true;

  For i := 1 to 10 do TabAxe[i] := TrouveArgument(S,'VALAXENIV'+(IntToStr(i)),'');

  i:=1;
  Title := 'Calendrier à la journée : ' + CodeSession;
  While TabAxe[i] <> '' do
  begin
    Title := Title + ' / ' + TabAxe[i];
    i:= i+1;
  end;

  Ecran.Caption := Title;
  THLabel(GetControl('SESSION')).Caption:=Title;


  { EVI / Redimensionnement de la colonne date }
  GridCalend.ColWidths[1] := 130;

  //ARRONDI
  GridCalend.ColFormats[2]:='#,####0.0000';
  GridCalend.ColAligns[2]:=taRightJustify;
  GridCalend.ColFormats[3]:='#,####0.0000';
  GridCalend.ColAligns[3]:=taRightJustify;
  GridCalend.ColTypes[2]:='F';

  GridCalend.ColEditables[1]:=false;
  GridCalend.ColEditables[2]:=false;
  GridCalend.ColEditables[5]:=false;

  GridCalend.ColEditables[3] := true;
  GridCalend.ColEditables[4] := true;
  GridCalend.ColEditables[6] := true;

  GridCalend.OnCellEnter:=CellEnter;
  GridCalend.OnCellExit:=CellExit;

  GridCalend.OnClick := CellClick;
  GridCalend.OnSetEditText:=CellEdit;

  GridCalend.OnColumnWidthsChanged := Resize;

  GridCalend.GetCellCanvas:=GetCellCanvas;
  GridCalend.OnRowEnter:=RowEnter;

  TToolBarButton97(getcontrol('BTNSOLDE')).OnClick := BtnSolde;
  TToolBarButton97(getcontrol('BIMPRIMER')).OnClick := BImprimerOnClick;


  GridCalend.ColColors[1]:=clBlack;

  { EVI / Récupération des dates de la session }
  Q:=MOPenSql('SELECT QBS_DATEDEBC,QBS_DATEFINC,QBS_DATEDEBREF,QBS_DATEFINREF,QBS_VUEARBRE FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+CodeSession+'"',
             'QUFVBPCALENDJOURM_TOF (OnArgument).',true);
  if not Q.eof then
  begin
    DateDebC := Q.Fields[0].AsDateTime;
    DateFinC := Q.Fields[1].AsDateTime;
    DateDebRef := Q.Fields[2].AsDateTime;
    DateFinRef := Q.Fields[3].AsDateTime;
    ValAff:=Q.fields[4].asString;
  end;
  Ferme(Q);

  THLabel(GetControl('DATEDEB')).Caption:=DateTimeToStr(DateDebC);
  THLabel(GetControl('DATEFIN')).Caption:=DateTimeToStr(DateFinC);

  DateDebSession := DateDebC;
  DateFinSession := DateFinC;

  { EVI / Formatage par semaine (conforme au stockage dans la base) }
  NumSemDeb:=NumSemaine(DateDebC,AnSemDeb);
  DateDebSem:=PremierJourSemaine(NumSemDeb,AnSemDeb);

  //Correction : Ajustement du décalage au lundi le plus proche
  case BPInitialise of
    2 : DateDebSemR := DateDebSem;
  else
    begin
      if DateDebSession-DateDebSem < 4 then DateDebSemR := DateDebSem - 7
      else DateDebSemR := DateDebSem;
    end;
  end;

  NumSemFin:=NumSemaine(DateFinC,AnSemFin);
  If NumSemfin=52 then DateFinSem:=PremierJourSemaine(1,AnSemFin+1)-1
  else DateFinSem:=PremierJourSemaine(NumSemFin+1,AnSemFin)-1;

         { EVI / Donne un résultat erroné avec Session Semaine 38719 - 3880
             if (NumSemFin>=NumSemDeb)
             then nbSem:=NumSemFin - NumSemDeb + 1 + (AnSemFin-AnSemDeb)*52
             else nbSem:=(52 - NumSemDeb + 1) + NumSemFin + 1 + (AnSemFin-AnSemDeb)*52; }

  nbSem:=NbSemaineIntervalle(DateDebSemR,DateFinSem);

  //Test si la session contient une année bissextile
  For i:= AnSemDeb to AnSemFin do
  begin
    if IsLeapYear(i) then
    begin
      { EVI / Appel année (aD) -> année courante (i) IsLeapYear }
       if DateAppartIntervalle(encodedate(i,2,29),DateDebC,DateFinC)
       then LeapFeb := LeapFeb + 1;
    end;
  end;

  GridCalend.RowCount:=(7*nbSem);
  SetLength(TabVerifPrct,nbSem+1);
  SetLength(TabNumPeriod,nbSem+1);
  SetLength(TabSumPeriodPrevu,nbSem+1);

  //création de la tob
  laTob := Tob.Create('_QT', nil, -1);
  TobUpdate := Tob.Create('QBPDETCALENDREP', nil, -1);

  for i:=0 to (7*nbSem)-1 do
  begin
    Tf := Tob.Create('_QGRF', latob, -1);
    Tf.AddChampSupValeur('JOUR', '');
    Tf.AddChampSupValeur('PRCTHISTO', '');
    Tf.AddChampSupValeur('PRCTPREVU', '');
    Tf.AddChampSupValeur('COM', '');
    Tf.AddChampSupValeur('HISTO', '');
    Tf.AddChampSupValeur('PREVU', '');
    Tf.AddChampSupValeur('PRCTGLOBAL','');
  end;

  { EVI / ListMailleJour pour correspondance HISTO - PREVU : VALEURI = 2 = Semaine }
  ListMailleJour:=TListMaille.create();
  InitialiseListeMaille(VALEURI('2'),DateDebC,DateFinC,DateDebRef,DateFinRef,ListMailleJour);

  { EVI / Correspondance entre DateCourante et DateReference pour la période appellée }
  Mi:=TMaille(ListMailleJour[0]);
  DateCourante := Mi.DateDebReference;
  DatePrevu := Mi.DateDebCourante;

  //Calcul du décalage
  case BPInitialise of
    2 : Decalage:=0;
  else
    begin
      DecalageR := Abs(Round(Mi.DateDebReference-DateDebRef));
      DecalageC := Abs(Round(DateDebSem-DateDebC));
      Decalage := DecalageC - DecalageR;

      //Correction : Ajustement du décalage au lundi le plus proche
      if DateDebSession-DateDebSem < 4 then
      begin
        if Decalage >=0 then Decalage := 7 - Decalage
        else Decalage := 7 + Decalage
      end;

    end;
  end; //CASE

  { EVI / Récupération du détail du calendrier }
  If ModeFiche = 'CALENDCONSULT' then
  begin
    for j:=1 to nbSem do
    begin
      GridCalend.ColColors[1]:=clblack;
      Q:=MOPenSql('SELECT QBE_PRCTJOUR1,QBE_PRCTJOUR2,QBE_PRCTJOUR3,QBE_PRCTJOUR4,'+
                  'QBE_PRCTJOUR5,QBE_PRCTJOUR6,QBE_PRCTJOUR7,'+
                  'QBE_COMMENTAIRE1,QBE_COMMENTAIRE2,'+
                  'QBE_COMMENTAIRE3,QBE_COMMENTAIRE4,QBE_COMMENTAIRE5,'+
                  'QBE_COMMENTAIRE6,QBE_COMMENTAIRE7,'+
                  'QBE_PRCTHISTO1,QBE_PRCTHISTO2,QBE_PRCTHISTO3,QBE_PRCTHISTO4,'+
                  'QBE_PRCTHISTO5,QBE_PRCTHISTO6,QBE_PRCTHISTO7'+
                  ' FROM QBPDETCALENDREP '+
                  'WHERE QBE_CALENDREP="'+CodeSession+
                  '" AND QBE_VALAXENIV1="'+TabAxe[1]+'" AND QBE_VALAXENIV2="'+TabAxe[2]+
                  '" AND QBE_VALAXENIV3="'+TabAxe[3]+'" AND QBE_VALAXENIV4="'+TabAxe[4]+
                  '" AND QBE_VALAXENIV5="'+TabAxe[5]+'" AND QBE_VALAXENIV6="'+TabAxe[6]+
                  '" AND QBE_VALAXENIV7="'+TabAxe[7]+'" AND QBE_VALAXENIV8="'+TabAxe[8]+
                  '" AND QBE_VALAXENIV9="'+TabAxe[9]+'" AND QBE_VALAXENIV10="'+TabAxe[10]+
                  '" AND QBE_DATEDEBSEM="'+USDATETIME(DateDebSemR+((j-1)*7))+'" ',
                 'QUFVBPCALENDJOURM_TOF (OnArgument).',true);
      if not Q.eof
      then
      begin
        for i:=1 to 7 do
        begin
          laTob.Detail[i+((j-1)*7)-1].putvalue('PRCTPREVU',Q.fields[i-1].asFloat);
          laTob.Detail[i+((j-1)*7)-1].putvalue('COM',Q.fields[i-1+7].asString);
          laTob.Detail[i+((j-1)*7)-1].putvalue('PRCTGLOBAL',Q.fields[i-1+14].asFloat);
        end;
      end;
      ferme(Q);
    end;
  end;

  { EVI / Récupération du calendrier type }
  If ModeFiche = 'CALENDTYPE' then
  begin
    Q:=MOpenSql('SELECT QBN_PRCT1,QBN_PRCT2,QBN_PRCT3,QBN_PRCT4,'+
                'QBN_PRCT5,QBN_PRCT6,QBN_PRCT7 FROM QBPCALENDREP '+
                'WHERE QBN_CALENDREP="'+Calendrier+'"',
                'QUFVBPCALENDJOURM_TOF (OnArgument).',true);
    if not Q.eof
    then
    begin
      for j:=1 to nbSem do
      begin
        for i:=1 to 7 do
        begin
          laTob.Detail[i+((j-1)*7)-1].putvalue('PRCTPREVU',Q.fields[i-1].asFloat);
          laTob.Detail[i+((j-1)*7)-1].putvalue('COM','');
        end;
      end;
    end;
    ferme(Q);
  end;

  { EVI / Si il ne s'agit pas d'une consultation, on récupère les dates des pièces dans QBPCUBETMP }
  If ModeFiche <> 'CALENDCONSULT' then
  begin
    //Récupération des axes
    Case ContextBP of
      0,1 : begin //Mode-GC
            if ValAff = 'PTC' then SQLValAff:='QBQ_CAHISTO,QBQ_CAPREVU';              //CA TTC net
            if ValAff = 'QTE' then SQLValAff:='QBQ_HISTO,QBQ_PREVU';                //Quantité
            if ValAff = 'PHT' then SQLValAff:='QBQ_CAHISTO2,QBQ_CAPREVU2';             //CA HT net
            if ValAff = 'UHT' then SQLValAff:='QBQ_CAHISTO3,QBQ_CAPREVU3';             //CA HT avant remise
            if ValAff = 'UTC' then SQLValAff:='QBQ_CAHISTO4,QBQ_CAPREVU4';             //CA TTC avant remise
            if ValAff = 'PAH' then SQLValAff:='QBQ_CAHISTO5,QBQ_CAPREVU5';             //Cumul achat PMAP
            if ValAff = 'MAR' then SQLValAff:='QBQ_CAHISTO6,QBQ_CAPREVU6';             //Marge
          end;
      2 : begin
            if ValAff = 'DC1' then SQLValAff:='QBQ_CAHISTO,QBQ_CAPREVU';
            if ValAff = 'CD1' then SQLValAff:='QBQ_CAHISTO2,QBQ_CAPREVU2';
          end;
    end; //CASE

    i:=1;
    AxeCube := '';
    While TabAxe[i] <> '' do
    begin
      if TabAxe[i]= '[---]' then AxeCube := AxeCube + ' AND (QBQ_VALAXECT'+IntToStr(i)+' is null OR QBQ_VALAXECT'+IntToStr(i)+'="")'
      else AxeCube := AxeCube + ' AND QBQ_VALAXECT'+IntToStr(i)+'="'+TabAxe[i]+'"';
      i:= i+1;
    end;

    { EVI / Requête sur la table QBPCUBETMP : récupération des dates, valeurs pour les pièces }
    Q:=MOpenSql('SELECT QBQ_DATEJOUR,'+SQLValAff+' FROM QBPCUBETMP WHERE QBQ_CODESESSION="'+CodeSession+'"'+AxeCube+
                ' AND QBQ_DATEJOUR >= "'+USDateTime(DateCourante)+'" ORDER BY QBQ_DATEJOUR','QUFVBPCALENDJOURM_TOF (OnArgument).',true);

    if not Q.eof then
    begin
    { EVI / Remplissage TabHisto }
      for j:=1 to nbSem do
      begin
        for i:=1 to 7 do
        begin
          DateCt := Q.Fields[0].AsDateTime;

          SommeHisto := 0;
          SommePrevu := 0;
          while DateCourante = DateCt do
          begin
            SommeHisto := SommeHisto + Q.fields[1].asFloat;
            Q.Next;
            DateCt := Q.Fields[0].AsDateTime;
            if Q.eof then break;
          end;

          while DatePrevu = DateCt do
          begin
            SommePrevu := SommePrevu + Q.fields[2].asFloat;
            Q.Next;
            DateCt := Q.Fields[0].AsDateTime;
            if Q.eof then break;
          end;

          laTob.Detail[i+((j-1)*7)-1].putvalue('HISTO',SommeHisto);
          laTob.Detail[i+((j-1)*7)-1].putvalue('PREVU',SommePrevu);
          DateCourante := DateCourante+1;
          DatePrevu := DatePrevu+1;
        end;
      end;
    end;
    ferme(Q);
  end
  else
  begin
    { EVI / Calcul de la table historique dans le cas d'une consultation }
    for i:=1 to nbSem*7 do
    begin
      if SommeHistoArbre <> 0 then laTob.Detail[i-1].putvalue('HISTO',(laTob.Detail[i-1].getvalue('PRCTGLOBAL')*SommeHistoArbre/100))
      else
      begin
        laTob.Detail[i-1].putvalue('HISTO','0,00');
        if SommePrevuArbre <> 0 then laTob.Detail[i-1].putvalue('PREVU',(laTob.Detail[i-1].getvalue('PRCTGLOBAL')*SommePrevuArbre/100))
      end;
    end;
  end;
end ;


procedure TOF_QUFVBPCALENDJOURM.CellClick(Sender: TObject);
var G:THGrid ;
    ColEnter:integer;
    jour:string;
    ValeurJour:TDateTime;
begin
  G:=THGrid(Sender) ;
  LigneEnter:=G.row;
  ColEnter:=G.col;

  jour:=copy(GridCalend.CellValues[1,G.Row],13,11);
  ValeurJour:=StrToDateTime(trim(copy(GridCalend.CellValues[1,G.Row],13,11)));

  if ((ColEnter = 3) AND ((DateDebSession<=Valeurjour)  AND  (Valeurjour<=DateFinSession))) then
  begin
   TPopUpMenu(GetControl('POPUPARBRE')).OnPopUp := PopupArbre;
   if RadioGrpMode.ItemIndex = 0 then
   TPopUpMenu(GetControl('POPUPARBRE')).Items[0].Caption := TraduireMemoire('Répartir le reliquat sur le reste de la période')
   else TPopUpMenu(GetControl('POPUPARBRE')).Items[0].Caption := TraduireMemoire('Répartir le reliquat sur les dates non sélectionnées');
   TPopUpMenu(GetControl('POPUPARBRE')).AutoPopup := true
  end
  else TPopUpMenu(GetControl('POPUPARBRE')).AutoPopup := false;
end;

procedure TOF_QUFVBPCALENDJOURM.PopupArbre(Sender: TObject);
begin
  TMenuItem(GetControl('RECALCRELIQ')).OnClick := RecalcReliq_OnClick;
end;

procedure  TOF_QUFVBPCALENDJOURM.RecalcReliq_OnClick(Sender: TObject);
var ValeurJour : TDateTime;
    Reliquat,ValCellRecalc,MaxValue,VerifPrct : Double;
    i,imax,Ind,NbLineSelected,FirstLineSelected : integer;
    TabLineSelected,TabIndiceSelected : array of integer;
    LineSelected: boolean;
begin
  nbJour := 0;
  VerifPrct := 0;
  MaxValue:=0;
  iMax:=0;
  LigneEnter:=GridCalend.row;

  Reliquat := THNumEdit(getcontrol('EDTSOLDE')).Value;

  if RadioGrpMode.ItemIndex = 0 then
  begin
    ValeurJour:=StrToDateTime(trim(copy(GridCalend.CellValues[1,LigneEnter],13,11)));
    SetControlText('TH_PERIODE',copy(GridCalend.CellValues[1,LigneEnter],1,7)) ;

    DonneDetailPeriod(ValeurJour,0);

    Reliquat:=Reliquat / nbJourPeriod ;

    for i := 0 to nbJourPeriod do
    begin
      if (DateDebSession<=ValeurJour)  AND  (ValeurJour<=DateFinSession) then
      begin
        if nbJour <> i then
        begin
          laTob.GetLigneGrid(GridCalend,LigneEnter - nbJour + i,';JOUR;PRCTHISTO;PRCTPREVU;COM');
          ValCellRecalc := StrToFloat(GridCalend.CellValues[3,LigneEnter - nbJour + i]) + Reliquat;
          laTob.SetDouble('PRCTPREVU',ValCellRecalc);
          laTob.PutLigneGrid(GridCalend,LigneEnter - nbJour + i,false,false,'JOUR;PRCTHISTO;PRCTPREVU;COM');
          if MaxValue < ValCellRecalc then
          begin
            MaxValue := ValCellRecalc;
            iMax:=i;
          end;
        end;
        VerifPrct:=VerifPrct+StrToFloat(GridCalend.CellValues[3,LigneEnter - nbJour + i])
      end;
    end;

    //Rectification --> 100%
    if (Arrondi(VerifPrct,4)<>100.0000) AND (Arrondi(VerifPrct,4)<>0.0000) then
    begin
      laTob.GetLigneGrid(GridCalend,LigneEnter - nbJour + iMax,';JOUR;PRCTHISTO;PRCTPREVU;COM');
      ValCellRecalc := StrToFloat(GridCalend.CellValues[3,LigneEnter - nbJour + iMax]) + (100-VerifPrct);
      laTob.SetDouble('PRCTPREVU',ValCellRecalc);
      laTob.PutLigneGrid(GridCalend,LigneEnter - nbJour + iMax,false,false,'JOUR;PRCTHISTO;PRCTPREVU;COM');
    end;

    RecalculTot(false,true,0);
    RecalculTotPeriode ;

  end
  else
  begin
    NbLineSelected := 0;
    if GridCalend.nbSelected <> 0 then
    begin
      //Contrôle si les lignes sélectionnées appartiennent à la même période
      GridCalend.GotoLeBookMark(0);
      FirstLineSelected := GridCalend.Row;
      ValeurJour:=StrToDateTime(trim(copy(GridCalend.CellValues[1,FirstLineSelected],13,11)));
      DonneDetailPeriod(ValeurJour,0);

      for Ind := 0 to GridCalend.nbSelected - 1 do
      begin
        GridCalend.GotoLeBookMark(Ind);
        if GridCalend.Row <= (FirstLineSelected - nbJour + nbJourPeriod) then NbLineSelected := Ind
        else
        begin
          NbLineSelected := Ind -1 ;
          break;
        end;
      end;

      SetLength(TabLineSelected,NbLineSelected+1);
      SetLength(TabIndiceSelected,NbLineSelected+1);
      for Ind := 0 to NbLineSelected do
      begin
        GridCalend.GotoLeBookMark( Ind );
        TabLineSelected[Ind]:=GridCalend.Row;
      end;

      ValeurJour:=StrToDateTime(trim(copy(GridCalend.CellValues[1,TabLineSelected[0]],13,11)));
      SetControlText('TH_PERIODE',copy(GridCalend.CellValues[1,TabLineSelected[0]],1,7)) ;

      DonneDetailPeriod(ValeurJour,0);

      Reliquat:=Reliquat / (nbJourPeriod-NbLineSelected) ;

      for Ind := 0 to NbLineSelected do
      begin
        TabIndiceSelected[Ind]:=TabLineSelected[Ind]-(TabLineSelected[0]-nbJour);
      end;

      for i := 0 to nbJourPeriod do
      begin
        LineSelected := false;
        for Ind := 0 to NbLineSelected do
        begin
          if TabIndiceSelected[Ind] = i then LineSelected := true;
        end;
        if LineSelected = false then
        begin
          laTob.GetLigneGrid(GridCalend,TabLineSelected[0] - nbJour + i,';JOUR;PRCTHISTO;PRCTPREVU;COM');
          ValCellRecalc := StrToFloat(GridCalend.CellValues[3,TabLineSelected[0] - nbJour + i]) + Reliquat;
          laTob.SetDouble('PRCTPREVU',ValCellRecalc);
          laTob.PutLigneGrid(GridCalend,TabLineSelected[0] - nbJour + i,false,false,'JOUR;PRCTHISTO;PRCTPREVU;COM');
          if MaxValue < ValCellRecalc then
          begin
            MaxValue := ValCellRecalc;
            iMax:=i;
          end;
        end;
        VerifPrct:=VerifPrct+StrToFloat(GridCalend.CellValues[3,TabLineSelected[0] - nbJour + i])
      end;

      //Rectification --> 100%
      if (Arrondi(VerifPrct,4)<>100.0000) AND (Arrondi(VerifPrct,4)<>0.0000) then
      begin
        laTob.GetLigneGrid(GridCalend,TabLineSelected[0] - nbJour + iMax,';JOUR;PRCTHISTO;PRCTPREVU;COM');
        ValCellRecalc := StrToFloat(GridCalend.CellValues[3,TabLineSelected[0] - nbJour + iMax]) + (100-VerifPrct);
        laTob.SetDouble('PRCTPREVU',ValCellRecalc);
        laTob.PutLigneGrid(GridCalend,TabLineSelected[0] - nbJour + iMax,false,false,'JOUR;PRCTHISTO;PRCTPREVU;COM');
      end;

      GridCalend.ClearSelected;

      RecalculTot(false,true,0);
      RecalculTotPeriode ;
    end
    else HShowMessage('1;Répartition impossible;Aucune ligne n''est sélectionnée.;W;O;O;O;','','');
  end;

end;

procedure TOF_QUFVBPCALENDJOURM.OnClose ;
var rep : integer;
begin
  Inherited ;
  if ValidationAuto<>'TRUE' then
  begin
    if not OkValid
    then
    begin
      LastError:=-1;
      OkValid:=true;
    end
    else
    begin
      if StateFiche then
      begin
        if ModeFiche = 'CALENDCONSULT' then
        begin
          rep:=HShowmessage('1;Calendrier.;Confirmez-vous l''annulation des mofications en cours ?;Q;YN;N;N', '', '');
          if rep=7 then LastError:=-1 else LastError:=0;
        end
        else
        begin
          rep:=HShowmessage('1;Calendrier.;Le calendrier n''a pas été enregistré. Voulez-vous continuer ?;Q;YN;N;N', '', '');
          if rep=7 then LastError:=-1 else LastError:=0;
        end
      end else LastError:=0;
    end
  end
end ;

procedure TOF_QUFVBPCALENDJOURM.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPCALENDJOURM.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPCALENDJOURM.BImprimerOnClick(Sender: TObject);
begin
  LanceEtatTOB('E','QUC','QU1',LaTob,True,False,False,TPageControl(GetControl('PCEtat')),'',Ecran.Caption,False)
 end ;

procedure TOF_QUFVBPCALENDJOURM.BtnSolde(Sender: TObject);
var somme,valueprevu:double;
    i:integer;
begin
  laTob.GetLigneGrid(GridCalend,LigneEnter,';JOUR;PRCTHISTO;PRCTPREVU;COM');
  valueprevu:=laTob.getValue('PRCTPREVU');
  laTob.SetDouble('PRCTPREVU',valueprevu+THNumEdit(getcontrol('EDTSOLDE')).Value);
  laTob.PutLigneGrid(GridCalend,LigneEnter,false,false,'JOUR;PRCTHISTO;PRCTPREVU;COM');
  somme:=0;
  for i := 1 to laTob.Detail.Count do
  begin
    laTob.GetLigneGrid(GridCalend,i,';JOUR;PRCTHISTO;PRCTPREVU;COM');
    if Latob.getString('PRCTPREVU') = '' then valueprevu:=0 else valueprevu:=laTob.getValue('PRCTPREVU');
    somme:=somme+valueprevu;
  end;
  RecalculTot(false,true,0);
  RecalculTotPeriode ;
end;

procedure TOF_QUFVBPCALENDJOURM.GetCellCanvas(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var ValeurJour:TDateTime;
    numPer,anPer:integer;
    anQ,moisQ,QuinzQ,NumQuinzQ:word;
    DecalColor : integer;

    function VerifLeap(Date:TDateTime): Boolean;
    var Yr, Mnth, Day: Word;
    begin
      DecodeDate(Date, Yr, Mnth, Day);
      if ((IsLeapYear(Yr)) AND (Mnth >= 2)) then result := true
      else result := false;
    end;

begin
  if (ARow > 0) and (ACol > 1) and (ACol<>4) then
  begin
    ValeurJour:=StrToDateTime(trim(copy(GridCalend.CellValues[1,ARow],13,11)));
    if ACol=5 then
    begin
      //Gestion des années bissextiles
      { EVI / ToDo session avec année bissextile > 1 }
      if LeapFeb > 0 then
      begin
        //Vérifie si l'année est bissextile et si le mois de fevrier est passé
        if VerifLeap(ValeurJour) then
        begin
          if Decalage >= 0 then DecalColor := Decalage + 1
          else DecalColor := Decalage - 1;
        end
        else DecalColor := Decalage;
      end
      else DecalColor := Decalage;

      if (((Valeurjour+DecalColor)<DateDebSession)  OR  (DateFinSession<(Valeurjour+DecalColor))) then
      begin
        Canvas.Font.Color:=clRed;
        Canvas.Brush.Color:=clBtnFace;
      end
      else
      begin
        case BPInitialise of
        2 : //Semaine
          begin
            if Odd(NumSemaine(Valeurjour)) then Canvas.Font.Color:=clActiveCaption
            else  Canvas.Font.Color:=clBlack;
          end;
        3 :  //Quinzaine
          begin
            Date_AnNumQuinz((Valeurjour+DecalColor),anQ,moisQ,QuinzQ,NumQuinzQ);
            if Odd(NumQuinzQ) then Canvas.Font.Color:=clActiveCaption
            else Canvas.Font.Color:=clBlack;
          end;
        4 : //Mois
          begin
            if Odd(NumMois((Valeurjour+DecalColor))) then Canvas.Font.Color:=clActiveCaption
            else Canvas.Font.Color:=clBlack;
          end;
        6 : //Trimestre
          begin
            DonneNumTrimestre((Valeurjour+DecalColor),numPer,anPer);
            if Odd(NumPer) then Canvas.Font.Color:=clActiveCaption
            else Canvas.Font.Color:=clBlack;
          end;
        7 : //Quadrimestre
          begin
            DonneNumQuadrimestre((Valeurjour+DecalColor),numPer,anPer);
            if Odd(NumPer) then Canvas.Font.Color:=clActiveCaption
            else Canvas.Font.Color:=clBlack;
          end;
        else Canvas.Font.Color:=clBlack; //CASE  { EVI / ToDo : Mois 4-4-5 }
        end; //CASE
        Canvas.Brush.Color:=clWindow;
      end;
    end
    else
    begin

      if ((Valeurjour<DateDebSession)  OR  (DateFinSession<Valeurjour)) then
      begin
        Canvas.Font.Color:=clRed;
        Canvas.Brush.Color:=clBtnFace;
      end
      else
      begin
        case BPInitialise of
        2 : //Semaine
          begin
            if Odd(NumSemaine(Valeurjour)) then Canvas.Font.Color:=clActiveCaption
            else  Canvas.Font.Color:=clBlack;
          end;
        3 :  //Quinzaine
          begin
            Date_AnNumQuinz(ValeurJour,anQ,moisQ,QuinzQ,NumQuinzQ);
            if Odd(NumQuinzQ) then Canvas.Font.Color:=clActiveCaption
            else Canvas.Font.Color:=clBlack;
          end;
        4 : //Mois
          begin
            if Odd(NumMois(Valeurjour)) then Canvas.Font.Color:=clActiveCaption
            else Canvas.Font.Color:=clBlack;
          end;
        6 : //Trimestre
          begin
            DonneNumTrimestre(Valeurjour,numPer,anPer);
            if Odd(NumPer) then Canvas.Font.Color:=clActiveCaption
            else Canvas.Font.Color:=clBlack;
          end;
        7 : //Quadrimestre
          begin
            DonneNumQuadrimestre(Valeurjour,numPer,anPer);
            if Odd(NumPer) then Canvas.Font.Color:=clActiveCaption
            else Canvas.Font.Color:=clBlack;
          end;
        else Canvas.Font.Color:=clBlack; //CASE  { EVI / ToDo : Mois 4-4-5 }
        end; //CASE
        Canvas.Brush.Color:=clWindow;
      end;
    end;
  end
  else
  begin
    if ACol=1 then
    begin
      Canvas.Brush.Color:=clBtnFace;
      Canvas.Font.Color:=clBlack;
    end;
  end;
end;

procedure TOF_QUFVBPCALENDJOURM.RecalculTot(Verif,Prct:Boolean;ligne:integer);
var ValCellPrevu,PeriodBug,PerioDispatch: string;
    FormulPrevu,SommePrevu,SommeVerifPrct:double ;
    jour,ValeurJour:TDateTime;
    i,numPeriodVerif,numPerioDispatch:Integer;
    TabDebPeriod,TabFinPeriod : array of integer;
begin
  SommeVerifPrct:=0;
  SommePrevu:=0;
  DateFinPeriod:=iDate1900;
  FormulPrevu:=0;
  numPeriod:=1;
  numPeriodVerif := 0;
  IncPeriod:=true;
  nbJour:=0;

  SetLength(TabDebPeriod,nbSem+1);
  SetLength(TabFinPeriod,nbSem+1);  //A Utiliser ?

  for i:=1 to GridCalend.RowCount-1 do
  begin
    jour:=StrToDateTime(trim(copy(GridCalend.CellValues[1,i],13,11)));
    if (DateDebSession<=jour)  AND  (jour<=DateFinSession) then
    begin
      if IncPeriod=true then DonneDetailPeriod(jour,i);

      case BPInitialise of
      1,5 :   { EVI / ToDo : Mois 4-4-5 }
        begin
          DateFinPeriod:=jour;
          FormulPrevu:=SommePrevuArbre;
        end;
      2,3,4,6,7 : FormulPrevu:=TabSumPeriodPrevu[numPeriod];
      end; //CASE

      if not (jour=DateFinPeriod) then
      begin
        //Stockage de la période et des dates bornes de la période
        if numPeriodVerif <> numPeriod then
        begin
          numPeriodVerif := NumPeriod;
          TabDebPeriod[numPeriodVerif] := i-1;
          TabFinPeriod[numPeriodVerif-1] := i-2;
        end;

        if Prct=true then
        begin
          if GridCalend.CellValues[3,i]='' then latob.Detail[i-1].putValue('PRCTPREVU',0)
          else latob.Detail[i-1].putValue('PRCTPREVU',StrToFloat(GridCalend.CellValues[3,i]));
          latob.Detail[i-1].putValue('PREVU',Valeur(GridCalend.Cells[3,i])*FormulPrevu/100);
          SommeVerifPrct:=SommeVerifPrct+latob.detail[i-1].GetValue('PRCTPREVU');
          SommePrevu := SommePrevu + latob.detail[i-1].GetValue('PREVU');
          latob.Detail[i-1].putValue('COM',GridCalend.CellValues[4,i]);
        end
        else
        begin
          ValCellPrevu := GridCalend.Cells[6,i];
          if ValCellPrevu = '' then ValCellPrevu := '0';
          ValCellPrevu:=StringReplace(ValCellPrevu,' ','',[rfReplaceAll, rfIgnoreCase]);
          if FormulPrevu <> 0 then latob.Detail[i-1].putValue('PRCTPREVU',StrToFloat(ValCellPrevu)*100/FormulPrevu)
          else latob.Detail[i-1].putValue('PRCTPREVU',0);
          latob.Detail[i-1].putValue('PREVU',StrToFloat(ValCellPrevu));
          SommeVerifPrct:=SommeVerifPrct+latob.detail[i-1].GetValue('PRCTPREVU');
          SommePrevu := SommePrevu + latob.detail[i-1].GetValue('PREVU');
          latob.Detail[i-1].putValue('COM',GridCalend.CellValues[4,i]);
        end;
      end
      else
      begin
        if Prct=true then
        begin
          if GridCalend.CellValues[3,i]='' then latob.Detail[i-1].putValue('PRCTPREVU',0)
          else latob.Detail[i-1].putValue('PRCTPREVU',StrToFloat(GridCalend.CellValues[3,i]));
          latob.Detail[i-1].putValue('PREVU',Valeur(GridCalend.Cells[3,i])*FormulPrevu/100);
          SommePrevu := SommePrevu + latob.detail[i-1].GetValue('PREVU');
          latob.Detail[i-1].putValue('COM',GridCalend.CellValues[4,i]);
          SommeVerifPrct:=SommeVerifPrct+latob.detail[i-1].GetValue('PRCTPREVU');
          TabNumPeriod[numPeriod]:=copy(GridCalend.CellValues[1,i],1,7);
          TabVerifPrct[numPeriod]:=SommeVerifPrct;
          if ((TabSumPeriodPrevu[numPeriod] <> 0) and (TabVerifPrct[numPeriod] = 0)) then
          begin
            if PeriodBug = '' then PeriodBug := IntToStr(numPeriod) else PeriodBug := PeriodBug + ';' + IntToStr(numPeriod);
          end;
          SommeVerifPrct:=0;
          IncPeriod:=true;
          numPeriod:=numPeriod+1;
        end
        else
        begin
          ValCellPrevu := GridCalend.Cells[6,i];
          if ValCellPrevu = '' then ValCellPrevu := '0';
          ValCellPrevu:=StringReplace(ValCellPrevu,' ','',[rfReplaceAll, rfIgnoreCase]);
          if FormulPrevu <> 0 then latob.Detail[i-1].putValue('PRCTPREVU',StrToFloat(ValCellPrevu)*100/FormulPrevu)
          else latob.Detail[i-1].putValue('PRCTPREVU',0);
          latob.Detail[i-1].putValue('PREVU',StrToFloat(ValCellPrevu));
          SommeVerifPrct:=SommeVerifPrct+latob.detail[i-1].GetValue('PRCTPREVU');
          SommePrevu := SommePrevu + latob.detail[i-1].GetValue('PREVU');
          latob.Detail[i-1].putValue('COM',GridCalend.CellValues[4,i]);
          TabNumPeriod[numPeriod]:=copy(GridCalend.CellValues[1,i],1,7);
          TabVerifPrct[numPeriod]:=SommeVerifPrct;
          SommeVerifPrct:=0;
          IncPeriod:=true;
          numPeriod:=numPeriod+1;
        end;
      end;
    end
    else
    begin
      latob.Detail[i-1].putValue('PRCTHISTO','');
      latob.Detail[i-1].putValue('PRCTPREVU','');
      latob.Detail[i-1].putValue('PREVU','');
      latob.Detail[i-1].putValue('COM','');
    end;
    if (((jour+decalage)<DateDebSession)  OR  (DateFinSession<(jour+decalage)))
    then latob.Detail[i-1].putValue('HISTO','');
  end;

  //Contrôle sur des pourcentages décalés qui n'ont pu être affectée, affectation à la première date de la période
  if Verif then
  begin
    if PeriodBug <> '' then
    begin
      PerioDispatch   := ReadTokenSt(PeriodBug);
      While PerioDispatch <> '' do
      begin
        NumPerioDispatch := StrToInt(PerioDispatch);
        latob.Detail[TabDebPeriod[NumPerioDispatch]].putValue('PRCTPREVU',100);
        latob.Detail[TabDebPeriod[NumPerioDispatch]].putValue('PREVU',TabSumPeriodPrevu[NumPerioDispatch]);
        SommePrevu := SommePrevu + TabSumPeriodPrevu[NumPerioDispatch];
        PerioDispatch   := ReadTokenSt(PeriodBug);
      end;
    end;
  end;

  THNumEdit(getcontrol('EDTPREV')).Value:=SommePrevu;

  if CellValid=true then
  begin
    if Ligne=0 then Ligne := LigneEnter;
    ValeurJour:=StrToDateTime(trim(copy(GridCalend.CellValues[1,Ligne],13,11)));
    DonneDetailPeriod(ValeurJour,0);
  end;

  laTob.PutGridDetail(GridCalend,false,false,';PRCTHISTO;PRCTPREVU;COM;HISTO;PREVU');

end;

procedure TOF_QUFVBPCALENDJOURM.RecalculTotPeriode;
var ValeurJour : TDateTime;
    Total : Double;
    i : integer;
begin
  Total := 0;
  nbJour := 0;
  LigneEnter:=GridCalend.row;

  ValeurJour:=StrToDateTime(trim(copy(GridCalend.CellValues[1,LigneEnter],13,11)));
  SetControlText('TH_PERIODE',copy(GridCalend.CellValues[1,LigneEnter],1,7)) ;

  DonneDetailPeriod(ValeurJour,0);

  for i := 0 to nbJourPeriod do
  begin
    if (DateDebSession<=ValeurJour)  AND  (ValeurJour<=DateFinSession) then
    begin
      ValCell := GridCalend.CellValues[3,LigneEnter - nbJour + i];
      Total := Total + StrToFloat(ValCell);
    end;
  end;

  THNumEdit(getcontrol('EDTTOTAL1')).Value:=Total;
  THNumEdit(getcontrol('EDTSOLDE')).Value:=100-Total;

end ;

procedure TOF_QUFVBPCALENDJOURM.CellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  RecalculTotPeriode ;
  CellValid:=false;
end;

procedure TOF_QUFVBPCALENDJOURM.CellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var ValeurJour:TDateTime;
begin
  if CellValid=true then
  begin
    if ((ACol=3) AND (ValCell = '')) then
    begin
      GridCalend.CellValues[3,LigneEnter] := '0,0000';
//      MAJTableGrid;
      laTob.GetLigneGrid(GridCalend,LigneEnter,';JOUR;PRCTHISTO;PRCTPREVU;COM');
    end;
    If ACol=3 then RecalculTot(false,true,ARow) else RecalculTot(false,false,ARow);
  end;

  if RadioGrpMode.ItemIndex = 1 then
  begin
    ValeurJour:=StrToDateTime(trim(copy(GridCalend.CellValues[1,ARow],13,11)));
    if ((Valeurjour<DateDebSession)  OR  (DateFinSession<Valeurjour)) then
    begin
      if GridCalend.IsSelected(ARow) = true then GridCalend.FlipSelection(ARow) ;
    end;
  end;

end;

procedure TOF_QUFVBPCALENDJOURM.CellEdit(Sender: TObject; ACol, ARow: Integer; const Value: hString);
begin
  CellValid:=true;
  ValCell:=Value;
  StateFiche := true;
end;

procedure TOF_QUFVBPCALENDJOURM.RowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
var ValeurJour:TDateTime;
begin
  ValeurJour:=StrToDateTime(trim(copy(GridCalend.CellValues[1,ou],13,11)));
  if (Valeurjour<DateDebSession)  OR  (DateFinSession<Valeurjour)
  then
  begin
    GridCalend.ColEditables[3] := False;
    GridCalend.ColEditables[6] := False
  end
  else
  begin
    GridCalend.ColEditables[3] := true;
    GridCalend.ColEditables[6] := true;
  end;
end;

procedure TOF_QUFVBPCALENDJOURM.Resize( Sender : TObject );
  function CalculPos (Fin:integer):integer;
  var i,Position:integer;
  begin
    Position:=0;
    for i:= 0 to Fin do Position:=Position+GridCalend.colWidths[i];
    Result:=Position;
  end;
begin
  THLabel(getcontrol('TOTAL')).Left := 20 + GridCalend.colWidths[1];
  TToolBarButton97(getcontrol('BTNSOLDE')).Left := 20 + GridCalend.colWidths[1];
  THLabel(getcontrol('TH_PERIODE')).Left := GridCalend.colWidths[1] + 60;
  THLabel(getcontrol('TOTAL1')).Left := GridCalend.colWidths[1] + 50;

  THNumEDIT(getcontrol('EDTTOTAL1')).Width := GridCalend.colWidths[3];
  THNumEDIT(getcontrol('EDTSOLDE')).Width := GridCalend.colWidths[3];
  THNumEDIT(getcontrol('EDTHISTO')).Width := GridCalend.colWidths[5];
  THNumEDIT(getcontrol('EDTPREV')).Width := GridCalend.colWidths[6];
  THEDIT(getcontrol('EDTPREVORIG')).Width := GridCalend.colWidths[6];

  THNumEDIT(getcontrol('EDTTOTAL1')).Left:=CalculPos(2);
  THNumEDIT(getcontrol('EDTSOLDE')).Left:=CalculPos(2);
  THNumEDIT(getcontrol('EDTHISTO')).Left:=CalculPos(4);
  THNumEDIT(getcontrol('EDTPREV')).Left:=CalculPos(5);
  THEDIT(getcontrol('EDTPREVORIG')).Left:=CalculPos(5);

  RadioGrpMode.Width :=GridCalend.colWidths[1];
  RadioGrpMode.Left := CalculPos(0);
end;

procedure TOF_QUFVBPCALENDJOURM.RadioGrpModeState(Sender: TObject);
begin
   if RadioGrpMode.ItemIndex = 0 then GridCalend.Options := GridCalend.Options + [goEditing]
   else GridCalend.Options := GridCalend.Options - [goEditing];
end;

Initialization
  registerclasses ( [ TOF_QUFVBPCALENDJOURM ] ) ;
end.
