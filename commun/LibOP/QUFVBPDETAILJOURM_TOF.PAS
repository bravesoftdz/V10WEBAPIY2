{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 21/07/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFVBPDETAILJOURM ()
Mots clefs ... : TOF;QUFVBPDETAILJOURM
*****************************************************************}
Unit QUFVBPDETAILJOURM_TOF ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     uTob,
     sysutils,
     HCtrls,
     HEnt1,
     UTOF,
     graphics,
     grids,
     BPMaille,
     BPBasic,
     Math,
     HPanel,
     controls,
     ExtCtrls;

Type
  TOF_QUFVBPDETAILJOURM = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

private
     CodeSession:hString;
     LeapFeb,nbSemX,BPInitialise,LigneEnter,ColEnter:integer;
     val,SommeHisto,SommePrevu:double;
     DateDebSemX,DateDebX,DateFinX,DateFinRecalcX,DateDebSem,DateDebSession,DateFinSession,DateDebRef,DateFinRef:TDateTime;
     GridCalend:THGrid;
     Resized,OKCalend,Recalc:boolean;
     Mi:TMaille;
     laTob:Tob;
     TitlePanel:THPanel;
     procedure GetCellCanvas( ACol, ARow : Longint; Canvas : TCanvas; AState : TGridDrawState );
     procedure RemplitGrid;
     procedure RecalculTot;
     procedure Resize( Sender : TObject );
     procedure DrawCell;
     procedure PopupCalend(Sender: TObject);
     procedure CellClick(Sender: TObject);
     procedure ChooseDate_OnClick(Sender: TObject);
     procedure BImprimerOnClick(Sender: TObject);
  end ;

Implementation

uses {$IFDEF EAGLCLIENT}
      UtileAGL,
     {$ELSE}
      EdtREtat,
     {$ENDIF}
      HTB97,HmsgBox,UUtil,windows,menus,forms,ComCtrls,BPUtil, Types;

function NumMois( Date : TDateTime ) : integer;
var Annee, Mois, Jour : Word;
begin
  DecodeDate( Date, Annee, Mois, Jour );
  Result := Mois;
end;

//Année bissextile + Mois de fevrier ou supérieur
function VerifLeap(Date:TDateTime): Boolean;
var Yr, Mnth, Day: Word;
begin
  DecodeDate(Date, Yr, Mnth, Day);
  if ((IsLeapYear(Yr)) AND (Mnth >= 2)) then result := true
  else result := false;
end;

procedure TOF_QUFVBPDETAILJOURM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPDETAILJOURM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPDETAILJOURM.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPDETAILJOURM.OnLoad ;
begin
  Inherited ;
  RemplitGrid;
  RecalculTot;
  Resize(self);
end ;

procedure TOF_QUFVBPDETAILJOURM.RemplitGrid;
var i,j,k,numSem,numPer,anPer:integer;
    anQ,moisQ,QuinzQ,NumQuinzQ:word;
    days: array[1..7] of hString;
    jour,JourRefHS:TDateTime;
begin
  days[1] := TraduireMemoire('Di');
  days[2] := TraduireMemoire('Lu');
  days[3] := TraduireMemoire('Ma');
  days[4] := TraduireMemoire('Me');
  days[5] := TraduireMemoire('Je');
  days[6] := TraduireMemoire('Ve');
  days[7] := TraduireMemoire('Sa');

  k:=0;

  for j:=1 to nbSemX*7 do
  begin
    i:=j;
    i:=i-k;
    jour:=PLUSDATE(DateDebSemX,j-1,'J');

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
      else //CASE { EVI / ToDo : Mois 445 }
        begin
          numSem:=NumSemaine(jour);
          Latob.Detail[i-1].putValue('JOUR',format(TraduireMemoire('Sem  ')+'%s : %-2s %s ',[MetZero(IntToStr(numSem),2),days[DayOfWeek(jour)],DateTimeToStr(jour)]));
        end;
    end;  //CASE

    JourRefHS := Mi.DateFinReference + (jour-Mi.DateFinCourante);
    //Correction : Ajustement du décalage au lundi le plus proche
    if BPInitialise > 2 then
    begin
      if DateDebSession-DateDebSem < 4 then JourRefHS := JourRefHS + 7;
    end;
    //Gestion des années bissextiles
    { EVI / ToDo : session avec année bissextile > 1 }
    if LeapFeb > 0 then
    begin
      //Vérifie si l'année est bissextile et si le mois de fevrier est passé
      if VerifLeap(jour) then JourRefHS := JourRefHS + 1;
    end;

    if ((DateDebRef+(DateDebX-DateDebSession)<=JourRefHS)  AND  (JourRefHS<=(DateDebRef+(DateFinX-DateDebSession)))) then
    begin
      SommeHisto:=SommeHisto+Latob.Detail[i-1].getValue('MTTHISTORIQUE');
      SommePrevu:=SommePrevu+Latob.Detail[i-1].getValue('MTTOBJECTIF');
    end
    else
    begin
      if ((Jour < DateDebX)  OR  (Jour > DateFinX )) then
      begin
        Latob.Detail[i-1].Free;
        k:=k+1;
      end;
    end;

  end;

  GridCalend.RowCount:= laTob.detail.count;

  laTob.PutGridDetail(GridCalend,false,false,'JOUR;MTTREALISE;MTTCUMREALISE;MTTOBJECTIF;DELTAOBJECTIF;PRCTOBJECTIF;MTTCUMOBJECTIF;DELTACUMOBJECTIF;PRCTCUMOBJECTIF;MTTHISTORIQUE;DELTAHISTORIQUE;PRCTHISTORIQUE;MTTCUMHISTORIQUE;DELTACUMHISTORIQUE;PRCTCUMHISTORIQUE');
end;

procedure TOF_QUFVBPDETAILJOURM.OnArgument (S : String ) ;
var nbaxe,LastAxe,numPer,anPer,i,j,AnSemDeb,NumSemDeb,NumSemDebX,AnSemDebX,NumSemFinX,AnSemFinX:integer;
    DateRealise,DateFinSemX,DateRefDebX,DateCurDebX,DateDebPeriod,DateFinPeriod,DateCt,DateDebC,DateFinC:TDateTime;
    NomCalend,AxeCalend,AxeCube,ValAff,SQLValAff,Title:hString;
    SommeHistoSQL,SommePrevuSQL,SommeRealiseSQL:double;
    TabAxe,TabAxeTemp:array[1..10] of hString;
    anQ,moisQ,QuinzQ,NumQuinzQ:word;
    ListMailleJour:TListMaille;
    FirstOccurence: boolean;
    Q:TQuery;
    Tf:Tob;
begin
  Inherited ;
  DateDebC:=0;
  DateFinC:=0;
  DateDebRef:=0;
  DateFinRef:=0;
  GridCalend:=THGrid(getcontrol('GRIDLOI'));
  TitlePanel := THPanel(getcontrol('TITLE_PANEL'));
  DateDebX:=StrToDateTime(TrouveArgument(S,'DELAI',''));
  CodeSession :=TrouveArgument(S,'SESSION','');
  ValAff := TrouveArgument(S,'VALAFF','');
  { EVI / ToDo : Val inutilisée }
  Val:=VALEUR(TrouveArgument(S,'VALEUR','0'));
  BPInitialise:=StrToInt(SessionBPInitialise(CodeSession));
  TToolBarButton97(getcontrol('BIMPRIMER')).OnClick := BImprimerOnClick;

  For i := 1 to 10 do
  begin
    if TrouveArgument(S,'VALAXENIV'+(IntToStr(i)),'null') = '' then TabAxe[i] := '[---]'
    else TabAxe[i] := TrouveArgument(S,'VALAXENIV'+(IntToStr(i)),'null')
  end;

  { EVI / Redimensionnement de la colonne date }
  GridCalend.ColWidths[0] := 130;
  GridCalend.ColAligns[0]:=taLeftJustify;

  GridCalend.ColEditables[0]:=false;
  for i := 1 to 14 do
  begin
    GridCalend.ColFormats[i]:='#,####0.00';
    GridCalend.ColAligns[i]:=taRightJustify;
    GridCalend.ColTypes[i]:='F';
    GridCalend.ColEditables[i]:=false;
  end;

  GridCalend.OnClick := CellClick;
  GridCalend.OnColumnWidthsChanged := Resize;

  GridCalend.GetCellCanvas:=GetCellCanvas;

  GridCalend.ColColors[1]:=clBlack;

  //Récupération des dates de la session
  Q:=MOPenSql('SELECT QBS_DATEDEBC,QBS_DATEFINC,QBS_DATEDEBREF,QBS_DATEFINREF,QBS_VUEARBRE FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+CodeSession+'"',
             'QUFVBPDETAILJOURM_TOF (OnArgument).',true);
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

  { EVI / ListMailleJour pour correspondance HISTO - PREVU : VALEURI = 2 = Semaine }
  ListMailleJour:=TListMaille.create();
  InitialiseListeMaille(VALEURI('2'),DateDebC,DateFinC,DateDebRef,DateFinRef,ListMailleJour);
  { EVI / Correspondance entre DateCourante et DateReference pour la période appellée }
  Mi:=TMaille(ListMailleJour[0]);

  NumSemDebX:=NumSemaine(DateDebX,AnSemDebX);
  DateDebSemX:=PremierJourSemaine(NumSemDebX,AnSemDebX);

  //Calcul du décalage
  if BPInitialise > 2 then
  begin
    //Correction : Ajustement du décalage au lundi le plus proche
    if DateDebSession-DateDebSem < 4 then DateDebSemX := DateDebSemX - 7;
  end;

  { EVI / Calcul de la fin de la période }
  case BPInitialise of
    2 : //Semaine
      DateFinX := DateDebX + 6  ;
    3 : //Quinzaine
      begin
        Date_AnNumQuinz(DateDebX,anQ,moisQ,QuinzQ,NumQuinzQ);
        if QuinzQ=1 then DateFinX:=EncodeDate(AnQ,moisQ,15)
        else DateFinX:=FinDeMois(DateDebX);
      end;
    4 : //Mois
      DateFinX := FinDeMois(DateDebX);
    6 : //Trimestre
      begin
        DonneNumTrimestre(DateDebX,numPer,anPer);
        DonneDateDebFinTrimestre(numPer,anPer,DateDebPeriod,DateFinPeriod);
        DateFinX := DateFinPeriod;
      end;
    7 : //Quadrimestre
      begin
        DonneNumQuadrimestre(DateDebX,numPer,anPer);
        DonneDateDebFinQuadrimestre(numPer,anPer,DateDebPeriod,DateFinPeriod);
        DateFinX := DateFinPeriod;
      end;
    //else //CASE { EVI / ToDo : Mois 4-4-5 }
  end; // CASE

  NumSemFinX:=NumSemaine(DateFinX,AnSemFinX);
  If NumSemfinX=52 then DateFinSemX:=PremierJourSemaine(1,AnSemFinX+1)-1
  else DateFinSemX:=PremierJourSemaine(NumSemFinX+1,AnSemFinX)-1;

  //Test si la session contient une année bissextile
  For i:= AnSemDeb to AnSemFinX do
  begin
    if IsLeapYear(i) then
    begin
      { EVI / Appel année (aD) -> année courante (i) IsLeapYear }
       if DateAppartIntervalle(encodedate(i,2,29),DateDebC,DateFinC)
       then LeapFeb := LeapFeb + 1;
    end;
  end;


  nbSemX:=NbSemaineIntervalle(DateDebSemX,DateFinSemX);

  DateFinRecalcX:=DateFinX;

  //création de la tob
  laTob := Tob.Create('_QT', nil, -1);
  for i:=0 to (7*nbSemX)-1 do
  begin
    Tf := Tob.Create('_QGRF', latob, -1);
    Tf.AddChampSupValeur('JOUR', '');
    Tf.AddChampSupValeur('MTTREALISE', 0);
    Tf.AddChampSupValeur('MTTCUMREALISE', 0);
    Tf.AddChampSupValeur('MTTOBJECTIF', 0);
    Tf.AddChampSupValeur('DELTAOBJECTIF', 0);
    Tf.AddChampSupValeur('PRCTOBJECTIF', 0);
    Tf.AddChampSupValeur('MTTCUMOBJECTIF', 0);
    Tf.AddChampSupValeur('DELTACUMOBJECTIF', 0);
    Tf.AddChampSupValeur('PRCTCUMOBJECTIF', 0);
    Tf.AddChampSupValeur('MTTHISTORIQUE', 0);
    Tf.AddChampSupValeur('DELTAHISTORIQUE', 0);
    Tf.AddChampSupValeur('PRCTHISTORIQUE', 0);
    Tf.AddChampSupValeur('MTTCUMHISTORIQUE', 0);
    Tf.AddChampSupValeur('DELTACUMHISTORIQUE', 0);
    Tf.AddChampSupValeur('PRCTCUMHISTORIQUE', 0);
    Tf.AddChampSupValeur('CALENDRIER', 0);
//    Tf.AddChampSupValeur('COM', '');
  end;

  { EVI / Récupération du détail du calendrier }
  AxeCalend := '';
  nbaxe:=0;
  For i := 1 to 10 do
  begin
    TabAxeTemp[i]:=TabAxe[i];
    if TabAxeTemp[i] <> 'null' then nbaxe:=i;
  end;

  LastAxe:=nbaxe;
  while nbaxe <> 0 do
  begin
    For i := 1 to nbaxe do
    begin
      if TabAxeTemp[i] <> 'null' then AxeCalend := AxeCalend + ' AND QBE_VALAXENIV'+IntToStr(i)+'="'+TabAxeTemp[i]+'"';
    end;
    if not ExisteSQL('Select ##TOP 1## QBE_CALENDREP From QBPDETCALENDREP where QBE_CALENDREP="'+CodeSession+'"'+AxeCalend) then
    begin
      { EVI / Cherche un calendrier sur un axe supérieur : dernier axe mis à vide }
      TabAxeTemp[LastAxe]:='';
      LastAxe:=LastAxe-1;
      if LastAxe=0 then break;
      AxeCalend:='';
    end
    else break;
  end;
  if LastAxe = 0 then
  begin
    HShowmessage('1;Attention;Aucun calendrier n''a été défini pour cette sélection.;E;O;', '', '');
    OKCalend:=false;
  end
  else OKCalend:=true;


  Title := TraduireMemoire('Détail à la journée : ') + CodeSession;
  FirstOccurence := false;
  For i := 1 to 10 do
  begin
    if TabAxe[i] <> 'null' then FirstOccurence := true;
    if (FirstOccurence = true) AND (TabAxe[i] <> 'null') then Title := Title + ' / ' + TabAxe[i];
  end;
  Ecran.Caption := Title;
  THLabel(GetControl('SESSION')).Caption:=Title;

  if LastAxe <> 0 then
  begin
    NomCalend := TraduireMemoire('Calendrier : ') + CodeSession;
    for i:=1 to LastAxe do NomCalend := NomCalend + ' / ' + TabAxeTemp[i];
  end
  else NomCalend := TraduireMemoire('Aucun calendrier défini');
  THLabel(getcontrol('LabelCalend')).Caption := NomCalend;

  for j:=1 to nbSemX do
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
                'WHERE QBE_CALENDREP="'+CodeSession+'"'+AxeCalend+
                ' AND QBE_DATEDEBSEM="'+USDATETIME(DateDebSemX+((j-1)*7))+'" ',
               'QUFVBPDETAILJOURM_TOF (OnArgument).',true);
    if not Q.eof
    then
    begin
      for i:=1 to 7 do Latob.Detail[i+((j-1)*7)-1].putValue('CALENDRIER',Q.fields[i-1].asFloat);
    end;
    ferme(Q);
  end;

  { EVI / Récupération des axes }
  Case ContextBP of
    0,1 : begin //Mode-GC
          if ValAff = 'PTC' then SQLValAff:='QBQ_CAHISTO,QBQ_CAPREVU';
          if ValAff = 'QTE' then SQLValAff:='QBQ_HISTO,QBQ_PREVU';
          if ValAff = 'PHT' then SQLValAff:='QBQ_CAHISTO2,QBQ_CAPREVU2';
          if ValAff = 'UHT' then SQLValAff:='QBQ_CAHISTO3,QBQ_CAPREVU3';
          if ValAff = 'UTC' then SQLValAff:='QBQ_CAHISTO4,QBQ_CAPREVU4';
          if ValAff = 'PAH' then SQLValAff:='QBQ_CAHISTO5,QBQ_CAPREVU5';
          if ValAff = 'MAR' then SQLValAff:='QBQ_CAHISTO6,QBQ_CAPREVU6';
        end;
    2 : begin //Compta
          if ValAff = 'DC1' then SQLValAff:='QBQ_CAHISTO,QBQ_CAPREVU';
          if ValAff = 'CD1' then SQLValAff:='QBQ_CAHISTO2,QBQ_CAPREVU2';
        end;
  end; //CASE

  AxeCube := '';

  For i := 1 to 10 do
  begin
    if TabAxe[i] <> 'null' then
    begin
      if TabAxe[i] = '[---]'
      then AxeCube := AxeCube + ' AND ((QBQ_VALAXECT'+IntToStr(i)+'="") OR (QBQ_VALAXECT'+IntToStr(i)+' IS NULL))'
      else AxeCube := AxeCube + ' AND QBQ_VALAXECT'+IntToStr(i)+'="'+TabAxe[i]+'"'
    end;
  end;

  //Correction : Ajustement du décalage au lundi le plus proche
  if BPInitialise > 2 then
  begin
    if DateDebSession-DateDebSem < 4
    then DateRefDebX:=TMaille(ListMailleJour[round((DateDebSemX+7-DateDebSem)/7)]).DateDebReference
    else DateRefDebX:=TMaille(ListMailleJour[round((DateDebSemX-DateDebSem)/7)]).DateDebReference;
  end
  else DateRefDebX:=TMaille(ListMailleJour[round((DateDebSemX-DateDebSem)/7)]).DateDebReference;

  { EVI / Requête sur la table QBPCUBETMP : récupération des dates, valeurs pour les pièces }
  Q:=MOpenSql('SELECT QBQ_DATEJOUR,'+SQLValAff+' FROM QBPCUBETMP WHERE QBQ_CODESESSION="'+CodeSession+'"'+AxeCube+
              ' AND QBQ_DATEJOUR >= "'+USDateTime(DateRefDebX)+'" ORDER BY QBQ_DATEJOUR','QUFVBPDETAILJOURM_TOF (OnArgument).',true);


  DateCurDebX := DateDebSemX;

  { EVI / Remplissage Histo,Objectif }
  if not Q.eof then
  begin
    for j:=1 to nbSemX do
    begin
      for i:=1 to 7 do
      begin
        DateCt := Q.Fields[0].AsDateTime;

        SommeHistoSQL := 0; SommePrevuSQL := 0;
        while DateRefDebX = DateCt do
        begin
          SommeHistoSQL := SommeHistoSQL + Q.fields[1].asFloat;
          SommePrevuSQL := SommePrevuSQL + Q.fields[2].asFloat;
          Q.Next;
          DateCt := Q.Fields[0].AsDateTime;
          if Q.eof then break;
        end;

        while DateCurDebX = DateCt do
        begin
          SommePrevuSQL := SommePrevuSQL + Q.fields[2].asFloat;
          Q.Next;
          DateCt := Q.Fields[0].AsDateTime;
          if Q.eof then break;
        end;

        Latob.Detail[i+((j-1)*7)-1].putValue('MTTHISTORIQUE',SommeHistoSQL);
        Latob.Detail[i+((j-1)*7)-1].putValue('MTTOBJECTIF',SommePrevuSQL);
        DateRefDebX := DateRefDebX+1;
        DateCurDebX := DateCurDebX+1;
      end;
    end;
  end;
  ferme(Q);

  { EVI / Récupération des axes }
  Case ContextBP of
    0,1 : begin //Mode-GC
          if ValAff = 'PTC' then SQLValAff:='QBQ_CAREALISE';
          if ValAff = 'QTE' then SQLValAff:='QBQ_REALISE';
          if ValAff = 'PHT' then SQLValAff:='QBQ_CAREALISE2';
          if ValAff = 'UHT' then SQLValAff:='QBQ_CAREALISE3';
          if ValAff = 'UTC' then SQLValAff:='QBQ_CAREALISE4';
          if ValAff = 'PAH' then SQLValAff:='QBQ_CAREALISE5';
          if ValAff = 'MAR' then SQLValAff:='QBQ_CAREALISE6';
        end;
    2 : begin //Compta
          if ValAff = 'DC1' then SQLValAff:='QBQ_CAREALISE';
          if ValAff = 'CD1' then SQLValAff:='QBQ_CAREALISE2';
        end
  end;

  DateRealise:=DateDebSemX;

  { EVI / Requête sur la table QBPCUBETMP : récupération des dates, valeurs pour les pièces }
  Q:=MOpenSql('SELECT QBQ_DATEJOUR,'+SQLValAff+' FROM QBPCUBETMP WHERE QBQ_CODESESSION="'+CodeSession+'"'+AxeCube+
              ' AND QBQ_DATEJOUR >= "'+USDateTime(DateRealise)+'" ORDER BY QBQ_DATEJOUR','QUFVBPDETAILJOURM_TOF (OnArgument).',true);

  { EVI / Remplissage TabRealise }
  if not Q.eof then
  begin
    for j:=1 to nbSemX do
    begin
      for i:=1 to 7 do
      begin
        DateCt := Q.Fields[0].AsDateTime;

        SommeRealiseSQL := 0;
        while DateRealise = DateCt do
        begin
          SommeRealiseSQL := SommeRealiseSQL + Q.fields[1].asFloat;
          Q.Next;
          DateCt := Q.Fields[0].AsDateTime;
          if Q.eof then break;
        end;

        Latob.Detail[i+((j-1)*7)-1].putValue('MTTREALISE',SommeRealiseSQL);
        DateRealise := DateRealise+1;
      end;
    end;
  end;
  ferme(Q);

end ;

procedure TOF_QUFVBPDETAILJOURM.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPDETAILJOURM.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPDETAILJOURM.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPDETAILJOURM.RecalculTot;
var SumReal,SumObj,SumDeltaObj,SumPrctObj,SumHisto,SumDeltaHisto,SumPrctHisto:double ;
    i:Integer ;
    JourRefHS,jour:TDateTime;
begin
  SumReal:=0;
  SumObj:=0;
  SumDeltaObj:=0;
  SumHisto:=0;
  SumDeltaHisto:=0;

  for i:=1 to GridCalend.RowCount-1 do
  begin
    jour:=StrToDateTime(trim(copy(GridCalend.CellValues[0,i],13,11)));
    JourRefHS := Mi.DateFinReference + (jour-Mi.DateFinCourante);
    //Correction : Ajustement du décalage au lundi le plus proche
    if BPInitialise > 2 then
    begin
      if DateDebSession-DateDebSem < 4 then JourRefHS := JourRefHS + 7;
    end;
    //Gestion des années bissextiles
    { EVI / ToDo : session avec année bissextile > 1 }
    if LeapFeb > 0 then
    begin
      //Vérifie si l'année est bissextile et si le mois de fevrier est passé
      if VerifLeap(jour) then JourRefHS := JourRefHS + 1;
    end;

    if (DateDebX<=jour)  AND  (jour<=DateFinRecalcX) then
    begin
      if i=1 then Latob.Detail[i-1].putValue('MTTCUMREALISE',Latob.Detail[i-1].getValue('MTTREALISE'))
      else Latob.Detail[i-1].putValue('MTTCUMREALISE',Latob.Detail[i-1].getValue('MTTCUMREALISE')+Latob.Detail[i-1].getValue('MTTREALISE'));

      //Recalcul de la table Objectif en fonction du calendrier
      Latob.Detail[i-1].putValue('MTTOBJECTIF',Latob.Detail[i-1].getValue('CALENDRIER')*SommePrevu/100);

      Latob.Detail[i-1].putValue('DELTAOBJECTIF',Latob.Detail[i-1].getValue('MTTREALISE')-Latob.Detail[i-1].getValue('MTTOBJECTIF'));
      if i=1 then Latob.Detail[i-1].putValue('MTTCUMOBJECTIF',Latob.Detail[i-1].getValue('MTTOBJECTIF'))
      else Latob.Detail[i-1].putValue('MTTCUMOBJECTIF',Latob.Detail[i-2].getValue('MTTCUMOBJECTIF')+Latob.Detail[i-1].getValue('MTTOBJECTIF'));
      if i=1 then Latob.Detail[i-1].putValue('DELTACUMOBJECTIF',Latob.Detail[i-1].getValue('DELTAOBJECTIF'))
      else Latob.Detail[i-1].putValue('DELTACUMOBJECTIF',Latob.Detail[i-2].getValue('DELTACUMOBJECTIF')+Latob.Detail[i-1].getValue('DELTAOBJECTIF'));

      if Latob.Detail[i-1].getValue('MTTOBJECTIF')<>0
      then Latob.Detail[i-1].putValue('PRCTOBJECTIF',Latob.Detail[i-1].getValue('MTTREALISE')/Latob.Detail[i-1].getValue('MTTOBJECTIF')*100)
      else Latob.Detail[i-1].putValue('PRCTOBJECTIF',0);
      if Latob.Detail[i-1].getValue('MTTCUMOBJECTIF')<>0
      then Latob.Detail[i-1].putValue('PRCTCUMOBJECTIF',Latob.Detail[i-1].getValue('MTTCUMREALISE')/Latob.Detail[i-1].getValue('MTTCUMOBJECTIF')*100)
      else Latob.Detail[i-1].putValue('PRCTCUMOBJECTIF',0);

      SumReal := SumReal + Latob.Detail[i-1].getValue('MTTREALISE');

      SumObj := SumObj + Latob.Detail[i-1].getValue('MTTOBJECTIF');
      SumDeltaObj := SumDeltaObj + Latob.Detail[i-1].getValue('DELTAOBJECTIF');
    end
    else
    begin
      Latob.Detail[i-1].putValue('DELTAOBJECTIF',0);
      Latob.Detail[i-1].putValue('MTTCUMOBJECTIF',0);
      Latob.Detail[i-1].putValue('DELTACUMOBJECTIF',0);
      Latob.Detail[i-1].putValue('PRCTOBJECTIF',0);
      Latob.Detail[i-1].putValue('PRCTCUMOBJECTIF',0);
    end;

    if ((DateDebRef+(DateDebX-DateDebSession)<=JourRefHS)  AND  (JourRefHS<=(DateDebRef+(DateFinRecalcX-DateDebSession)))) then
    begin
      Latob.Detail[i-1].putValue('DELTAHISTORIQUE',Latob.Detail[i-1].getValue('MTTREALISE')-Latob.Detail[i-1].getValue('MTTHISTORIQUE'));
      if i=1 then  Latob.Detail[i-1].putValue('MTTCUMHISTORIQUE',Latob.Detail[i-1].getValue('MTTHISTORIQUE'))
      else  Latob.Detail[i-1].putValue('MTTCUMHISTORIQUE',Latob.Detail[i-2].getValue('MTTCUMHISTORIQUE')+Latob.Detail[i-1].getValue('MTTHISTORIQUE'));
      if i=1 then Latob.Detail[i-1].putValue('DELTACUMHISTORIQUE',Latob.Detail[i-1].getValue('DELTAHISTORIQUE'))
      else Latob.Detail[i-1].putValue('DELTACUMHISTORIQUE',Latob.Detail[i-2].getValue('DELTACUMHISTORIQUE')+Latob.Detail[i-1].getValue('DELTAHISTORIQUE'));

      if Latob.Detail[i-1].getValue('MTTHISTORIQUE')<>0
      then Latob.Detail[i-1].putValue('PRCTHISTORIQUE',Latob.Detail[i-1].getValue('MTTREALISE')/Latob.Detail[i-1].getValue('MTTHISTORIQUE')*100)
      else Latob.Detail[i-1].putValue('PRCTHISTORIQUE',0);
      if  Latob.Detail[i-1].getValue('MTTCUMHISTORIQUE')
      then Latob.Detail[i-1].putValue('PRCTCUMHISTORIQUE',Latob.Detail[i-1].getValue('MTTCUMREALISE')/ Latob.Detail[i-1].getValue('MTTCUMHISTORIQUE')*100)
      else Latob.Detail[i-1].putValue('PRCTCUMHISTORIQUE',0);

      SumHisto := SumHisto + Latob.Detail[i-1].getValue('MTTHISTORIQUE');
      SumDeltaHisto := SumDeltaHisto + Latob.Detail[i-1].getValue('DELTAHISTORIQUE');
    end
    else
    begin
      Latob.Detail[i-1].putValue('DELTAHISTORIQUE',0);
      Latob.Detail[i-1].putValue('MTTCUMHISTORIQUE',0);
      Latob.Detail[i-1].putValue('DELTACUMHISTORIQUE',0);
      Latob.Detail[i-1].putValue('PRCTHISTORIQUE',0);
      Latob.Detail[i-1].putValue('PRCTCUMHISTORIQUE',0);
    end;
  end;

  laTob.PutGridDetail(GridCalend,false,false,'JOUR;MTTREALISE;MTTCUMREALISE;MTTOBJECTIF;DELTAOBJECTIF;PRCTOBJECTIF;MTTCUMOBJECTIF;DELTACUMOBJECTIF;PRCTCUMOBJECTIF;MTTHISTORIQUE;DELTAHISTORIQUE;PRCTHISTORIQUE;MTTCUMHISTORIQUE;DELTACUMHISTORIQUE;PRCTCUMHISTORIQUE');

  if SumObj<>0 then SumPrctObj:=SumReal/SumObj*100 else SumPrctObj:=0;
  if SumHisto<> 0 then SumPrctHisto:=SumReal/SumHisto*100 else SumPrctHisto := 0;

  THNumEdit(getcontrol('EDT_MTTREAL')).Value:=SumReal;

  THNumEdit(getcontrol('EDT_MTTOBJ')).Value:=SumObj;
  THNumEdit(getcontrol('EDT_DELTAOBJ')).Value:=SumDeltaObj;
  THNumEdit(getcontrol('EDT_PRCTOBJ')).Value:=SumPrctObj;

  THNumEdit(getcontrol('EDT_MTTHIST')).Value:=SumHisto;
  THNumEdit(getcontrol('EDT_DELTAHIST')).Value:=SumDeltaHisto;
  THNumEdit(getcontrol('EDT_PRCTHIST')).Value:=SumPrctHisto;
end;

procedure TOF_QUFVBPDETAILJOURM.DrawCell;
var TShReal,TShObj,TShHisto : Tshape;
 THLabReal,THLabObj,THLabHisto : THLabel;
begin
  TShReal := TShape(GetControl('ENTETEREAL'));
  TShObj := TShape(GetControl('ENTETEPREVU'));
  TShHisto := TShape(GetControl('ENTETEHISTO'));
  THLabReal := THLabel(GetControl('LABELREAL'));
  THLabObj := THLabel(GetControl('LABELPREVU'));
  THLabHisto := THLabel(GetControl('LABELHISTO'));

  TShReal.Left :=  GridCalend.CellRect(1,0).Left ;
  TShReal.Width :=  GridCalend.CellRect(2,0).Right - GridCalend.CellRect(1,0).Left  ;
  TShReal.Height := TitlePanel.Height;
  TShReal.Brush.Color := clBtnFace;
  TShReal.Pen.Style := psSolid;
  TShReal.Pen.Width := 1;
  TShReal.Pen.Color := clBtnShadow;

  TShObj.Left :=  GridCalend.CellRect(3,0).Left ;
  TShObj.Width :=  GridCalend.CellRect(8,0).Right - GridCalend.CellRect(3,0).Left  ;
  TShObj.Height := TitlePanel.Height;
  TShObj.Brush.Color := clBtnFace;
  TShObj.Pen.Style := psSolid;
  TShObj.Pen.Width := 1;
  TShObj.Pen.Color := clBtnShadow;

  TShHisto.Left :=  GridCalend.CellRect(9,0).Left ;
  TShHisto.Width :=  GridCalend.CellRect(14,0).Right - GridCalend.CellRect(9,0).Left  ;
  TShHisto.Height := TitlePanel.Height;
  TShHisto.Brush.Color := clBtnFace;
  TShHisto.Pen.Style := psSolid;
  TShHisto.Pen.Width := 1;
  TShHisto.Pen.Color := clBtnShadow;


  THLabReal.Left := GridCalend.CellRect(1,0).Left ;
  THLabReal.Width :=  GridCalend.CellRect(2,0).Right - GridCalend.CellRect(1,0).Left  ;
  THLabReal.Height := TitlePanel.Height;
  THLabReal.Caption := TraduireMemoire('Réalisé');

  THLabObj.Left := GridCalend.CellRect(3,0).Left ;
  THLabObj.Width :=  GridCalend.CellRect(8,0).Right - GridCalend.CellRect(3,0).Left  ;
  THLabObj.Height := TitlePanel.Height;
  THLabObj.Caption := TraduireMemoire('Objectif');

  THLabHisto.Left := GridCalend.CellRect(9,0).Left ;
  THLabHisto.Width :=  GridCalend.CellRect(14,0).Right - GridCalend.CellRect(9,0).Left  ;
  THLabHisto.Height := TitlePanel.Height;
  THLabHisto.Caption := TraduireMemoire('Historique');
end;

procedure TOF_QUFVBPDETAILJOURM.GetCellCanvas(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var ValeurJour:TDateTime;
    JourRefHS:TDateTime;
begin
  if ((ARow > 0) AND (ACol > 0)) then
  begin
    if ARow = LigneEnter then Canvas.Brush.Color:=clInactiveCaptionText else Canvas.Brush.Color:=clWhite;
    ValeurJour:=StrToDateTime(trim(copy(GridCalend.CellValues[0,ARow],13,11)));
    if ACol < 9 then
    begin
      //Colonnes Réalisé + Prévu
      if ((Valeurjour<DateDebX)  OR  (DateFinX<Valeurjour)) then
      begin
        Canvas.Font.Color:=clBtnFace;                       //clRed;
        Canvas.Brush.Color:=clBtnFace;
      end
      else
      begin
        Canvas.Font.Color:=clBlack;
        Canvas.Brush.Color:= clWhite;

        //Recalcul à date sur Réalisé + Prévu
        if Valeurjour > DateFinRecalcX then
        begin
          Canvas.Brush.Color:= cl3DLight;
          Canvas.Font.Style := [fsItalic];
        end else
        begin
          Canvas.Brush.Color:=clWhite;
          if ACol > 2 then Canvas.Font.Color:=clBlue
          else Canvas.Font.Color:=clBlack;
        end;
      end
    end
    else
    begin
      //Colonnes Historique
      JourRefHS := Mi.DateFinReference + (ValeurJour-Mi.DateFinCourante);
      //Correction : Ajustement du décalage au lundi le plus proche
      if BPInitialise > 2 then
      begin
        if DateDebSession-DateDebSem < 4 then JourRefHS := JourRefHS + 7;
      end;
      //Gestion des années bissextiles
      { EVI / ToDo : session avec année bissextile > 1 }
      if LeapFeb > 0 then
      begin
        //Vérifie si l'année est bissextile et si le mois de fevrier est passé
        if VerifLeap(ValeurJour) then JourRefHS := JourRefHS + 1;
      end;

      if ((DateDebRef+(DateDebX-DateDebSession)<=JourRefHS)  AND  (JourRefHS<=(DateDebRef+(DateFinX-DateDebSession)))) then
      begin
        Canvas.Font.Color:=clBlack;
        Canvas.Brush.Color:=clWhite;

        //Recalcul à date sur Historique
//        if Valeurjour > DateFinRecalcX then
        if JourRefHS > (DateDebRef+(DateFinRecalcX-DateDebSession)) then
        begin
          Canvas.Brush.Color:= cl3DLight;
          Canvas.Font.Style := [fsItalic];
        end else
        begin
          Canvas.Brush.Color:=clWhite;
          Canvas.Font.Color:=clBlack;
        end;

      end
      else
      begin
        Canvas.Font.Color:=clBtnFace;     //clRed;
        Canvas.Brush.Color:=clBtnFace;
      end;
    end;
  end
  else
  begin
    //Colonne date
    Canvas.Font.Color:=clBlack;
    Canvas.Brush.Color:=clBtnFace;
  end;

  if (Resized=true) then
  begin
    DrawCell;
    Resized:=false;
  end;
end;

procedure TOF_QUFVBPDETAILJOURM.CellClick(Sender: TObject);
var
   G:THGrid ;
    jour:hString;
    ValeurJour:TDateTime;
begin
  G:=THGrid(Sender) ;
  LigneEnter:=G.row;
  ColEnter:=G.col;

  jour:=copy(GridCalend.CellValues[0,G.Row],13,11);
  ValeurJour:=StrToDateTime(trim(copy(GridCalend.CellValues[0,G.Row],13,11)));

  if ((ColEnter = 0) AND ((DateDebX<=Valeurjour)  AND  (Valeurjour<=DateFinX))) then
  begin
   TPopUpMenu(GetControl('POPUPCALEND')).OnPopUp := PopupCalend;
   TPopUpMenu(GetControl('POPUPCALEND')).Items[0].Caption := TraduireMemoire('Recalcul à partir du') + jour;
   TPopUpMenu(GetControl('POPUPCALEND')).AutoPopup := true
  end
  else TPopUpMenu(GetControl('POPUPCALEND')).AutoPopup := false;
end;

procedure TOF_QUFVBPDETAILJOURM.PopupCalend(Sender: TObject);
begin
  TMenuItem(GetControl('CHOOSEDATE')).OnClick := ChooseDate_OnClick;
end;

procedure  TOF_QUFVBPDETAILJOURM.ChooseDate_OnClick(Sender: TObject);
begin
  DateFinRecalcX:=StrToDateTime(trim(copy(GridCalend.CellValues[0,LigneEnter],13,11)));
  if DateFinRecalcX=DateFinX then Recalc:=false else Recalc:=true;
  RecalculTot;
end;

procedure TOF_QUFVBPDETAILJOURM.Resize( Sender : TObject );

  function CalculPos (Fin:integer):integer;
  var i,Position:integer;
  begin
    Position:=0;
    for i:= 0 to Fin do Position:=Position+GridCalend.colWidths[i];
    Result:=Position;
  end;

begin
  GridCalend.ColWidths[0] := 130;
  Resized:=true;

  THLabel(getcontrol('LABEL_REAL')).Left:=GridCalend.colWidths[0];
  THLabel(getcontrol('LABEL_REAL')).Width:=GridCalend.colWidths[1];
  THLabel(getcontrol('LABEL_MTTREAL')).Left:=CalculPos(1);
  THLabel(getcontrol('LABEL_MTTREAL')).Width:=GridCalend.colWidths[2];
  THNumEdit(getcontrol('EDT_MTTREAL')).Left:=CalculPos(1);
  THNumEdit(getcontrol('EDT_MTTREAL')).Width:=GridCalend.colWidths[2];

  THLabel(getcontrol('LABEL_OBJ')).Left:=CalculPos(4);
  THLabel(getcontrol('LABEL_OBJ')).Width:=GridCalend.colWidths[5];

  THLabel(getcontrol('LABEL_MTTOBJ')).Left:=CalculPos(5);
  THLabel(getcontrol('LABEL_MTTOBJ')).Width:=GridCalend.colWidths[6];
  THNumEdit(getcontrol('EDT_MTTOBJ')).Left:=CalculPos(5);
  THNumEdit(getcontrol('EDT_MTTOBJ')).Width:=GridCalend.colWidths[6];

  THLabel(getcontrol('LABEL_DELTAOBJ')).Left:=CalculPos(6);
  THLabel(getcontrol('LABEL_DELTAOBJ')).Width:=GridCalend.colWidths[7];
  THNumEdit(getcontrol('EDT_DELTAOBJ')).Left:=CalculPos(6);
  THNumEdit(getcontrol('EDT_DELTAOBJ')).Width:=GridCalend.colWidths[7];

  THLabel(getcontrol('LABEL_PRCTOBJ')).Left:=CalculPos(7);
  THLabel(getcontrol('LABEL_PRCTOBJ')).Width:=GridCalend.colWidths[8];
  THNumEdit(getcontrol('EDT_PRCTOBJ')).Left:=CalculPos(7);
  THNumEdit(getcontrol('EDT_PRCTOBJ')).Width:=GridCalend.colWidths[8];

  THLabel(getcontrol('LABEL_HISTO')).Width:=GridCalend.colWidths[11];
  THLabel(getcontrol('LABEL_HISTO')).Left:=CalculPos(10);

  THLabel(getcontrol('LABEL_MTTHISTO')).Left:=CalculPos(11);
  THLabel(getcontrol('LABEL_MTTHISTO')).Width:=GridCalend.colWidths[12];
  THNumEdit(getcontrol('EDT_MTTHIST')).Left:=CalculPos(11);
  THNumEdit(getcontrol('EDT_MTTHIST')).Width:=GridCalend.colWidths[12];

  THLabel(getcontrol('LABEL_DELTAHISTO')).Left:=CalculPos(12);
  THLabel(getcontrol('LABEL_DELTAHISTO')).Width:=GridCalend.colWidths[13];
  THNumEdit(getcontrol('EDT_DELTAHIST')).Left:=CalculPos(12);
  THNumEdit(getcontrol('EDT_DELTAHIST')).Width:=GridCalend.colWidths[13];

  THLabel(getcontrol('LABEL_PRCTHISTO')).Left:=CalculPos(13);
  THLabel(getcontrol('LABEL_PRCTHISTO')).Width:=GridCalend.colWidths[14];
  THNumEdit(getcontrol('EDT_PRCTHIST')).Left:=CalculPos(13);
  THNumEdit(getcontrol('EDT_PRCTHIST')).Width:=GridCalend.colWidths[14];
end;

procedure TOF_QUFVBPDETAILJOURM.BImprimerOnClick(Sender: TObject);
begin
  LanceEtatTOB('E','QUC','QU2',LaTob,True,False,False,TPageControl(GetControl('PCEtat')),'',Ecran.Caption,False)
end ;


Initialization
  registerclasses ( [ TOF_QUFVBPDETAILJOURM ] ) ;
end.
