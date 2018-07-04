{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 22/07/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFQBPBUDGET ()
Mots clefs ... : TOF;QUFQBPBUDGET
*****************************************************************}
Unit QUFQBPBUDGET_TOF ;

Interface

Uses StdCtrls, 
     Controls,
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     EdtRetat,
{$else}
     utilEagl,
{$ENDIF}
     uTob,
     sysutils,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF ,
     htb97,
     SynScriptBP,
     windows,
     ed_tools;

const
    TabValue : array[ 0..6 ] of string = (
    ( 'SUM(IIF((QBQ_NATVALAFF="Historique"),QBQ_VALAFF2,0)) AS HISTO2,SUM(IIF((QBQ_NATVALAFF="Prévu"),QBQ_VALAFF2,0)) AS PREVU2,'+
      'SUM(IIF((QBQ_NATVALAFF="Réalisé"),QBQ_VALAFF2,0)) AS REALISE2' ),
    ( 'SUM(IIF((QBQ_NATVALAFF="Historique"),QBQ_VALAFF1,0)) AS HISTO1,SUM(IIF((QBQ_NATVALAFF="Prévu"),QBQ_VALAFF1,0)) AS PREVU1,'+
      'SUM(IIF((QBQ_NATVALAFF="Réalisé"),QBQ_VALAFF1,0)) AS REALISE1' ),
    ( 'SUM(IIF((QBQ_NATVALAFF="Historique"),QBQ_VALAFF3,0)) AS HISTO3,SUM(IIF((QBQ_NATVALAFF="Prévu"),QBQ_VALAFF3,0)) AS PREVU3,'+
      'SUM(IIF((QBQ_NATVALAFF="Réalisé"),QBQ_VALAFF3,0)) AS REALISE3' ),
    ( 'SUM(IIF((QBQ_NATVALAFF="Historique"),QBQ_VALAFF4,0)) AS HISTO4,SUM(IIF((QBQ_NATVALAFF="Prévu"),QBQ_VALAFF4,0)) AS PREVU4,'+
      'SUM(IIF((QBQ_NATVALAFF="Réalisé"),QBQ_VALAFF4,0)) AS REALISE4' ),
    ( 'SUM(IIF((QBQ_NATVALAFF="Historique"),QBQ_VALAFF5,0)) AS HISTO5,SUM(IIF((QBQ_NATVALAFF="Prévu"),QBQ_VALAFF5,0)) AS PREVU5,'+
      'SUM(IIF((QBQ_NATVALAFF="Réalisé"),QBQ_VALAFF5,0)) AS REALISE5' ),
    ( 'SUM(IIF((QBQ_NATVALAFF="Historique"),QBQ_VALAFF6,0)) AS HISTO6,SUM(IIF((QBQ_NATVALAFF="Prévu"),QBQ_VALAFF6,0)) AS PREVU6,'+
      'SUM(IIF((QBQ_NATVALAFF="Réalisé"),QBQ_VALAFF6,0)) AS REALISE6' ),
    ( 'SUM(IIF((QBQ_NATVALAFF="Historique"),QBQ_VALAFF7,0)) AS HISTO7,SUM(IIF((QBQ_NATVALAFF="Prévu"),QBQ_VALAFF7,0)) AS PREVU7,'+
      'SUM(IIF((QBQ_NATVALAFF="Réalisé"),QBQ_VALAFF7,0)) AS REALISE7' )
    );

Type
 TOF_QUFQBPBUDGET = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BVALIDER_OnCLick(Sender: TObject);
    procedure RempliTobPourEtatGlobal(const codeSession:string;var laTob:TOB);
    { EVI / Récupération des critères - multiselection}
    procedure DonneFiltre;
    procedure FEtatOnClick( Sender : TObject );
    procedure FEtatOnChange( Sender : TObject );
    procedure DonneDetailPeriod(Jour:TDateTime;typeSession:integer);
    procedure OnNouvelleRecherche ( Sender : TObject);
    function CopyTOB(T : TOB) : integer;
    function TobEnLigne(T : TOB) : integer;
    function RemplitColonneTob(T : TOB) : integer;
    private
    Contexte        : string;
    edtsession      : string;
    codeFiltre      : string;
    TobQBQ,TobMultiValQBQ,TobCurrentCalend : Tob;
    ComboModeleEtat : TComboBox;
    BValider: TToolbarButton97;
    { EVI / Date Edition Debut, Date Edition Fin }
    edtDateDeb,edtDateFin,DateEdDeb,DateEdFin : TDateTime;
    DateDebRef,DateFinRef,DateDebPeriod,DateFinPeriod:TDateTime;
    NbValAff,numSem,nbJour, nbJourPeriod : integer;
    IncPeriod:boolean;
    Histo,Prevu,Realise : Double;
    TabLibVal : array of hString;
    InProgress : boolean;

  end ;

Implementation
{ EVI / ToDo : [Voir avec L.Meunier] Fonction GetArgumentString = wCommuns (Mode) & UtilPGI (GC) }
Uses wCommuns, UtilPGI,
     BpCubeTmp,BpFctSession,
     ComCtrls,BPMaille,BPBAsic,
     UUtil,Filtre,BPUtil,StrUtils,
     {$IFDEF PAIEGRH}EntPaie,{$ENDIF PAIEGRH}
     {$IFDEF EAGLCLIENT}eQRS1;{$ELSE}QRS1;{$ENDIF}

procedure TOF_QUFQBPBUDGET.OnArgument (S : String ) ;
begin
  Inherited ;
  Contexte:= GetArgumentString(S,'CONTEXTE');
  BValider := GetControl('BValider') as TToolbarButton97;
  Ecran.OnKeyDown := FormKeyDown;

  if Contexte='PERIODE' then SetControlProperty('CBCUMUL', 'VISIBLE', false);
  if Assigned(GetControl('BVALIDER')) then
    TToolBarButton97(GetControl('BVALIDER')).OnClick := BVALIDER_OnClick;
  ComboModeleEtat := TComboBox( GetControl( 'FEtat' ) );
  ComboModeleEtat.OnClick := FEtatOnClick;
  ComboModeleEtat.OnChange := FEtatOnChange;

  if ContextBp = 3 then
  begin
    TFQRS1(Ecran).NatureEtat := 'QUP';
    TCheckBox(GetControl('DETAILJOUR')).Visible := false;
    TCheckBox(GetControl('VOIRCUMULS')).Visible := false;
    TCheckBox(GetControl('CBCUMUL')).Visible := false;
  end;
end ;

procedure TOF_QUFQBPBUDGET.DonneDetailPeriod(Jour:TDateTime;typeSession:integer);
var AnSemDeb,numPer,anPer:integer;
    anQ,moisQ,QuinzQ,NumQuinzQ:word;
begin
  case TypeSession of
    2 : //Semaine
      begin
        numSem:=NumSemaine(jour,AnSemDeb);
        DateDebPeriod:=PremierJourSemaine(numSem,AnSemDeb);
        DateFinPeriod:=DateDebPeriod + 7;
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
        DateDebPeriod := jour;
        DateFinPeriod := jour;
      end;
  end; // CASE
end;

function TOF_QUFQBPBUDGET.CopyTOB(T : TOB) : integer;
var NewTob : TOB;
begin
  // Copie la TOB T dans une TOB fille de TOBGrid
  NewTob := TOB.Create('toto', TobCurrentCalend, -1);
  NewTob.Dupliquer(T, false, true);
  result:=0;
end;

function TOF_QUFQBPBUDGET.TobEnLigne(T : TOB) : integer;
var NewTob : TOB;
i : integer;
begin
  // Copie la TOB T dans une TOB fille de TOBGrid
  if not MoveCurProgressForm(TraduireMemoire('Traitement des données...')) then result := 1 else result := 0;
  for i := 1 to NbValAff do
  begin
    Histo:=0;Prevu:=0;Realise:=0;
    if i = 1 then
    begin
      if T.GetString('HISTO2') <> '' then Histo := T.GetDouble('HISTO2') else Histo:=0;
      if T.GetString('PREVU2') <> '' then Prevu := T.GetDouble('PREVU2') else Prevu:=0;
      if T.GetString('REALISE2') <> '' then Realise := T.GetDouble('REALISE2') else Realise:=0;
    end
    else if i = 2 then
    begin
      if T.GetString('HISTO1') <> '' then Histo := T.GetDouble('HISTO1') else Histo:=0;
      if T.GetString('PREVU1') <> '' then Prevu := T.GetDouble('PREVU1') else Prevu:=0;
      if T.GetString('REALISE1') <> '' then Realise := T.GetDouble('REALISE1') else Realise:=0;
    end
    else
    begin
      if T.GetString('HISTO'+IntToStr(i)) <> '' then Histo := T.GetDouble('HISTO'+IntToStr(i)) else Histo:=0;
      if T.GetString('PREVU'+IntToStr(i)) <> '' then Prevu := T.GetDouble('PREVU'+IntToStr(i)) else Prevu:=0;
      if T.GetString('REALISE'+IntToStr(i)) <> '' then Realise := T.GetDouble('REALISE'+IntToStr(i)) else Realise:=0;
    end;
    //Stockage des différentes valeurs les unes à la suite des autres : chaque tob renseignée avec HISTO2,PREVU2,REALISE2 + LIBELLE
    T.AddChampSupValeur('LIBELLE',TabLibVal[i]);
    //Stockage de toutes les valeurs dans chaque TOB : HISTO1, HISTO2, HISTO3... + LIBELLE1, LIBELLE2...
    //T.AddChampSupValeur('LIBELLE'+IntToStr(i),TabLibVal[i]);
    //CurrentValue := i;
    //Stockage des différentes valeurs les unes à la suite des autres : chaque tob renseignée avec HISTO2,PREVU2,REALISE2 + LIBELLE
    if  TobMultiValQBQ.ParcoursTraitement(['QBQ_CODESESSION','QBQ_DATECT','QBQ_VALAXECT1','QBQ_VALAXECT2','QBQ_VALAXECT3',
    'QBQ_VALAXECT4','QBQ_VALAXECT5','QBQ_VALAXECT6','QBQ_VALAXECT7','QBQ_VALAXECT8','QBQ_VALAXECT9','QBQ_VALAXECT10','LIBELLE'],
    [T.GetValue('QBQ_CODESESSION'),T.GetValue('QBQ_DATECT'),T.GetValue('QBQ_VALAXECT1'),T.GetValue('QBQ_VALAXECT2'),
    T.GetValue('QBQ_VALAXECT3'),T.GetValue('QBQ_VALAXECT4'),T.GetValue('QBQ_VALAXECT5'),T.GetValue('QBQ_VALAXECT6'),
    T.GetValue('QBQ_VALAXECT7'),T.GetValue('QBQ_VALAXECT8'),T.GetValue('QBQ_VALAXECT9'),T.GetValue('QBQ_VALAXECT10'),T.GetValue('LIBELLE')],
    true,RemplitColonneTob) = 0 then
    begin
      T.PutValue('HISTO2',Histo);
      T.PutValue('PREVU2',Prevu);
      T.PutValue('REALISE2',Realise);
      NewTob := TOB.Create('Tob Temp', TobMultiValQBQ, -1);
      NewTob.Dupliquer(T, false, true);
    end;
    //Stockage des différentes valeurs les unes à la suite des autres : chaque tob renseignée avec HISTO2,PREVU2,REALISE2 + LIBELLE
{    if  TobMultiValQBQ.ParcoursTraitement(['QBQ_CODESESSION','QBQ_DATECT','QBQ_VALAXECT1','QBQ_VALAXECT2','QBQ_VALAXECT3',
    'QBQ_VALAXECT4','QBQ_VALAXECT5','QBQ_VALAXECT6','QBQ_VALAXECT7','QBQ_VALAXECT8','QBQ_VALAXECT9','QBQ_VALAXECT10'],
    [T.GetValue('QBQ_CODESESSION'),T.GetValue('QBQ_DATECT'),T.GetValue('QBQ_VALAXECT1'),T.GetValue('QBQ_VALAXECT2'),
    T.GetValue('QBQ_VALAXECT3'),T.GetValue('QBQ_VALAXECT4'),T.GetValue('QBQ_VALAXECT5'),T.GetValue('QBQ_VALAXECT6'),
    T.GetValue('QBQ_VALAXECT7'),T.GetValue('QBQ_VALAXECT8'),T.GetValue('QBQ_VALAXECT9'),T.GetValue('QBQ_VALAXECT10')],
    true,RemplitColonneTob) = 0 then
    begin
      if i = 1 then
      begin
        //A Revoir Histo <> 0 : Empèche de retenir des valeurs nulles --> Suprimer le contrôle
        if Histo <> 0 then T.PutValue('HISTO2',Histo);
        if Prevu <> 0 then T.PutValue('PREVU2',Prevu);
        if Realise <> 0 then T.PutValue('REALISE2',Realise);
      end
      else if i = 2 then
      begin
        if Histo <> 0 then T.PutValue('HISTO1',Histo);
        if Prevu <> 0 then T.PutValue('PREVU1',Prevu);
        if Realise <> 0 then T.PutValue('REALISE1',Realise);
      end
      else
      begin
        if Histo <> 0 then T.PutValue('HISTO'+IntToStr(i),Histo);
        if Prevu <> 0 then T.PutValue('PREVU'+IntToStr(i),Prevu);
        if Realise <> 0 then T.PutValue('REALISE'+IntToStr(i),Realise);
      end;
      NewTob := TOB.Create('Tob Temp', TobMultiValQBQ, -1);
      NewTob.Dupliquer(T, false, true);
    end;   }
  end;
end;

function TOF_QUFQBPBUDGET.RemplitColonneTob(T : TOB) : integer;
begin
  //Stockage des différentes valeurs les unes à la suite des autres : chaque tob renseignée avec HISTO2,PREVU2,REALISE2 + LIBELLE
  if Histo <> 0 then T.PutValue('HISTO2',Histo);
  if Prevu <> 0 then T.PutValue('PREVU2',Prevu);
  if Realise <> 0 then T.PutValue('REALISE2',Realise);

  //Stockage de toutes les valeurs dans chaque TOB : HISTO1, HISTO2, HISTO3... + LIBELLE1, LIBELLE2...
{ if CurrentValue = 1 then
  begin
    if Histo <> 0 then T.PutValue('HISTO2',Histo);
    if Prevu <> 0 then T.PutValue('PREVU2',Prevu);
    if Realise <> 0 then T.PutValue('REALISE2',Realise);
  end
  else if CurrentValue = 2 then
  begin
    if Histo <> 0 then T.PutValue('HISTO1',Histo);
    if Prevu <> 0 then T.PutValue('PREVU1',Prevu);
    if Realise <> 0 then T.PutValue('REALISE1',Realise);
  end
  else
  begin
    if Histo <> 0 then T.PutValue('HISTO'+IntToStr(CurrentValue),Histo);
    if Prevu <> 0 then T.PutValue('PREVU'+IntToStr(CurrentValue),Prevu);
    if Realise <> 0 then T.PutValue('REALISE'+IntToStr(CurrentValue),Realise);
  end;}
  result:=0;
end;

procedure TOF_QUFQBPBUDGET.BVALIDER_OnCLick(Sender: TObject);
Var
  decalage : double;
  ListMaille:TListMaille;
  CodeEtat,ChpPrevu: hString;
  TobDateSup,TobCalend : TOB;
  TypeSession,i,j,ParcoursTobEtat,ParcoursCurrentCalend,ParcourJours:integer;
  nivMaxCalend,nivMaxSession,EnregOrig,RI:integer;
  FindValue,TraitementOK:boolean;
  SumPrevu, Prevu:double;
  DateDebPeriodVerif,DateCalend : TDateTime;
  Q:TQuery;
  ValueAxeCalend,ValueAxeEtat : array [1..9] of variant;
  VerifParcoursTobEtat : integer;
  NumSemDeb,AnSemDeb,BPInitialise : integer;
  DateDebSem : TDateTime;
  Mi:TMaille;
  ValueAxe,LimitEdition : hString;
begin
  VerifParcoursTobEtat := -1;
  DateFinPeriod := 0;
  nivMaxCalend := 0;
  EdtSession := GetControlText('EDTSESSION');

  if (ControleSession(Edtsession)=false) then
  begin
    PGIINFO('Le code session doit être renseigné correctement.');
    exit;
  end;

  { EVI / Contrôle pour le cas où la tablette est vide (session non initialisée) }
  if GetCheckBoxState ('DETAILJOUR') = cbChecked then
  begin
    if (GetControlText('DATEJOURDEB')='') or (GetControlText('DATEJOURFIN')='') then
     begin
       DateEdDeb := 0;
       DateEdFin := 0;
     end else
     begin
       DateEdDeb := StrToDate(GetControlText('DATEJOURDEB'));
       DateEdFin := StrToDate(GetControlText('DATEJOURFIN'));
     end;
  end else
  begin
    if (GetControlText('DATEDEB')='') or (GetControlText('DATEFIN')='') then
    begin
      DateEdDeb := 0;
      DateEdFin := 0;
    end else
    begin
      DateEdDeb := StrToDate(GetControlText('DATEDEB'));
      DateEdFin := StrToDate(GetControlText('DATEFIN'));
    end;
  end;

  TobQBQ:= Tob.Create('QBQ', nil, -1);
  TobMultiValQBQ := Tob.Create('Tob Pour Paie', nil, -1);

  Try {Charge la Tob d'édition}

    { EVI / Avec les dates d'edition la table est regénéré dans tous les cas }
    {if Contexte='PERIODE' then RempliTobPourEtatGlobal(EdtSession, TobQBQ);
    //La première fois la table QBQ est vide
    if (Assigned(TobQBQ)) and (TobQBQ.Detail.Count=0) then
    begin}

    ChercheDateDDateFPeriode(edtsession,edtDateDeb,edtDateFin,DateDebRef,DateFinRef);

    if GetCheckBoxState ('DETAILJOUR') = cbChecked then
    begin
      { EVI / Contrôle pour le cas où les dates d'edition sortent des dates de la session }
      if ((DateEdDeb < edtDateDeb) OR (DateEdDeb = 0)) then
      begin
        DateEdDeb := edtDateDeb;
        SetControlText('DATEJOURDEB',DateTimeToStr(DateEdDeb));
      end;
      if ((DateEdFin > edtDateFin) OR (DateEdFin = 0))then
      begin
        DateEdFin := edtDateFin;
        SetControlText('DATEJOURFIN',DateTimeToStr(DateEdFin));
      end;
    end
    else
    begin
      if DateEdDeb = 0 then DateEdDeb := edtDateDeb;
      if DateEdFin = 0 then DateEdFin := edtDateFin;
    end;

    nivMaxSession:=ChercheNivMax(EdtSession);
    //Récupère les restrictions pour la génération du cube
    for i:=1 to nivMaxSession do
    begin
       ValueAxe := GetControlText('EDTAXE'+IntToStr(i));
       ValueAxe:= AnsiReplaceText(ValueAxe, ';', '","');
       if LimitEdition = '' then LimitEdition := '("'+ValueAxe+'")'
       else LimitEdition := LimitEdition + '@@' + '("'+ValueAxe+'")'
    end;

    InitMoveProgressForm(nil,TraduireMemoire('Edition standard'),TraduireMemoire('Veuillez patienter'), NbDeMaille(edtsession,'InitOK',DateEdDeb,DateEdFin) + NbMajCoeff(edtsession), True, False);
    InProgress := true;

    { EVI / Ajout Date Edition Debut, Date Edition Fin }
    //RemplitTableQBPCubeTmpPgi(edtSession,edtDateDeb,edtDateFin,DateEdDeb,DateEdFin);
    { EVI / Correction : même en détail jour on prend les bornes DATEDEB-DATEFIN pour générer le cube }
    try
      { EVI / Génération du cube avec restriction utilisateur }
      RemplitTableQBPCubeTmpPgi(edtSession,LimitEdition,edtDateDeb,edtDateFin,DateEdDeb,DateEdFin);
    except
      FiniMoveProgressForm;
    end;

    TraitementOK := MoveCurProgressForm(TraduireMemoire('Génération...'));

    (* OT 25/09/2006 *)
    if TraitementOK then
    begin
      if GetCheckBoxState ('DETAILJOUR') = cbChecked then
      begin
        { EVI / Correction du calcul du décalage }
        ListMaille:=TListMaille.create();
        InitialiseListeMaille(VALEURI('2'),edtDateDeb,edtDateFin,DateDebRef,DateFinRef,ListMaille);
        Mi:=TMaille(ListMaille[0]);

        BPInitialise:=StrToInt(SessionBPInitialise(Edtsession));

        //Correction : Ajustement du décalage au lundi le plus proche
        NumSemDeb:=NumSemaine(edtDateDeb,AnSemDeb);
        DateDebSem:=PremierJourSemaine(NumSemDeb,AnSemDeb);

        decalage:=Mi.DateDebCourante-Mi.DateDebReference;

        //Correction : Ajustement du décalage au lundi le plus proche
        if BPInitialise > 2 then
        begin
          if edtDateDeb-DateDebSem < 4 then decalage := Abs(decalage - 7);
        end;

        freeListMaille(ListMaille);
        ExecuteSQL ('UPDATE QBPCUBETMP SET QBQ_DATEJOUR=QBQ_DATEJOUR+' + FloatToStr(decalage) + ' WHERE QBQ_DATEJOUR<"' +
                      USDATETIME(DateEdDeb) + '"');
        ExecuteSQL ('DELETE FROM QBPCUBETMP WHERE  QBQ_DATEJOUR>"' + USDATETIME(DateEdFin) + '" OR QBQ_DATEJOUR<"' + USDATETIME(DateEdDeb) + '"');
      end;

      if contexte='PERIODE' then RempliTobPourEtatGlobal(EdtSession, TobQBQ);

      if GetCheckBoxState ('DETAILJOUR') = cbChecked then
      begin

        RI:=DonneValeurAffiche(EdtSession);
        case RI of
           1 : ChpPrevu:='QBR_OP1';
           2 : ChpPrevu:='QBR_QTEC';
           3 : ChpPrevu:='QBR_OP2';
           4 : ChpPrevu:='QBR_OP3';
           5 : ChpPrevu:='QBR_OP4';
           6 : ChpPrevu:='QBR_OP5';
           7 : ChpPrevu:='QBR_OP6';
        end;

        TypeSession := StrToInt(SessionBPInitialise(EdtSession));

        TobQBQ.Detail.Sort('QBQ_VALAXECT1;QBQ_VALAXECT2;QBQ_VALAXECT3;QBQ_VALAXECT4;QBQ_VALAXECT5;QBQ_VALAXECT6;QBQ_VALAXECT7;QBQ_VALAXECT8;QBQ_VALAXECT9;QBQ_VALAXECT10;QBQ_DATEJOUR');

        FiniMoveProgressForm;

        InitMoveProgressForm(nil,TraduireMemoire('Edition standard'),TraduireMemoire('Veuillez patienter'), TobQBQ.Detail.Count, False, False);

        TobCalend:= Tob.Create('QBE', nil, -1);
        try
          TobCalend.LoadDetailDBFromSQL('QBPDETCALENDREP',
                                        'SELECT QBE_CALENDREP,QBE_DATEDEBSEM,'+
                                        'QBE_PRCTJOUR1,QBE_PRCTJOUR2,QBE_PRCTJOUR3,QBE_PRCTJOUR4,QBE_PRCTJOUR5,QBE_PRCTJOUR6,QBE_PRCTJOUR7,'+
                                        'QBE_PRCTHISTO1,QBE_PRCTHISTO2,QBE_PRCTHISTO3,QBE_PRCTHISTO4,QBE_PRCTHISTO5,QBE_PRCTHISTO6,QBE_PRCTHISTO7,'+
                                        'QBE_VALAXENIV1,QBE_VALAXENIV2,QBE_VALAXENIV3,QBE_VALAXENIV4,QBE_VALAXENIV5,QBE_VALAXENIV6,QBE_VALAXENIV7,'+
                                        'QBE_VALAXENIV8,QBE_VALAXENIV9 FROM QBPDETCALENDREP '+
                                        'WHERE QBE_CALENDREP="'+edtSession+'"');


          //Détermine l'axe maximum définit pour le calendrier
          TobCalend.Detail.Sort('QBE_VALAXENIV1;QBE_VALAXENIV2;QBE_VALAXENIV3;QBE_VALAXENIV4;QBE_VALAXENIV5;QBE_VALAXENIV6;QBE_VALAXENIV7;QBE_VALAXENIV8;QBE_VALAXENIV9;QBE_VALAXENIV10;QBE_DATEDEBSEM');

          for i:=0 to TobCalend.Detail.Count-1 do
          begin
            for j:=1 to 9 do
            begin
              if TobCalend.Detail[i].GetValue('QBE_VALAXENIV'+IntToStr(j))<>'' then nivMaxCalend := j
            end;
          end;

          //Existe-il un calendrier définit pour un axe inférieur
          if nivMaxSession >= nivMaxCalend then
          begin

            ParcoursTobEtat:=0;
            EnregOrig := TobQBQ.Detail.Count-1;
            while ParcoursTobEtat <= EnregOrig do
            begin
              TobCurrentCalend:= Tob.Create('QBE', nil, -1);
              try
                //Protection pour le cas où aucune correspondance n'a été trouvé entre le calendrier et le cube
                //if (VerifParcoursTobEtat > 0) AND (VerifParcoursTobEtat = ParcoursTobEtat) then
                if (VerifParcoursTobEtat = ParcoursTobEtat) then
                begin
                  ParcoursTobEtat := ParcoursTobEtat + 1;
                  MoveCurProgressForm('Analyse des calendriers...');
                end;
                VerifParcoursTobEtat := ParcoursTobEtat;
                if VerifParcoursTobEtat > EnregOrig then break;
                //Recopie les valeurs qui correspondent au niveau maximum du calendrier
                for i := 1 to nivmaxCalend do
                ValueAxeEtat[i] := TobQBQ.Detail[ParcoursTobEtat].GetValue('QBQ_VALAXECT'+IntToStr(i));
                for i := nivmaxCalend+1 to 9 do ValueAxeEtat[i] := '';
                TobCalend.ParcoursTraitement(['QBE_VALAXENIV1','QBE_VALAXENIV2','QBE_VALAXENIV3','QBE_VALAXENIV4','QBE_VALAXENIV5',
                'QBE_VALAXENIV6','QBE_VALAXENIV7','QBE_VALAXENIV8','QBE_VALAXENIV9'],
                [ValueAxeEtat[1],ValueAxeEtat[2],ValueAxeEtat[3],ValueAxeEtat[4],ValueAxeEtat[5],ValueAxeEtat[6],ValueAxeEtat[7],ValueAxeEtat[8],ValueAxeEtat[9]],true, copyTob);

                DateFinPeriod :=0;
                DateDebPeriodVerif := 0;
                SumPrevu := 0;

                //Aucun calendrier pour cet enregistrement, on passe au suivant
                if TobCurrentCalend.Detail.count = 0 then
                begin
                  ParcoursTobEtat := ParcoursTobEtat + 1;
                  MoveCurProgressForm('Analyse des calendriers...');
                end;

                For ParcoursCurrentCalend := 0 to TobCurrentCalend.Detail.count-1 do
                begin
                  for ParcourJours := 1 to 7 do
                  begin

                    //Il existe une nouvelle valeur définit dans le calendrier
                    Prevu := TobCurrentCalend.Detail[ParcoursCurrentCalend].GetValue('QBE_PRCTJOUR'+IntToStr(ParcourJours));
                    DateCalend := TobCurrentCalend.Detail[ParcoursCurrentCalend].GetValue('QBE_DATEDEBSEM') + ParcourJours-1;

                    //Si la Date du calendrier est supérieure à la date d'edition --> sortie
                    if DateCalend > DateEdFin then break;

                    if (DateEdDeb <= DateCalend) AND (DateCalend <= DateEdFin  ) then
                    begin
                      //Retourne la période qui correspond à la date du calendrier
                      DonneDetailPeriod(DateCalend,TypeSession);
                      if DateDebPeriodVerif <> DateDebPeriod then
                      begin
                        //Récupération du prévu pour la sélection d'axe
                        SumPrevu := 0;
                        DateDebPeriodVerif := DateDebPeriod;

                        for i := 1 to nivMaxSession do
                        ValueAxeCalend[i] := TobQBQ.Detail[ParcoursTobEtat].GetValue('QBQ_VALAXECT'+IntToStr(i));
                        for i := nivMaxSession+1 to 9 do ValueAxeCalend[i] := '';

                        Q:=OpenSql('SELECT '+ChpPrevu+' FROM QBPARBRE WHERE QBR_CODESESSION="'+EdtSession+
                                   '" AND QBR_VALAXENIV1="'+ValueAxeCalend[1]+
                                   '" AND QBR_VALAXENIV2="'+ValueAxeCalend[2]+
                                   '" AND QBR_VALAXENIV3="'+ValueAxeCalend[3]+
                                   '" AND QBR_VALAXENIV4="'+ValueAxeCalend[4]+
                                   '" AND QBR_VALAXENIV5="'+ValueAxeCalend[5]+
                                   '" AND QBR_VALAXENIV6="'+ValueAxeCalend[6]+
                                   '" AND QBR_VALAXENIV7="'+ValueAxeCalend[7]+
                                   '" AND QBR_VALAXENIV8="'+ValueAxeCalend[8]+
                                   '" AND QBR_VALAXENIV9="'+ValueAxeCalend[9]+
                                   '" AND QBR_VALEURAXE="'+DateTimeToStr(DateDebPeriod)+'"',true);

                        if not Q.eof  then SumPrevu:=Q.fields[0].asFloat;
                        ferme(Q);
                      end;

                      //Si la valeur n'est pas nulle on continue
                      if SumPrevu <> 0 then
                      begin

                        //On parcours la tob : vérification si des valeurs correspondent
                        //Test uniquement jusqu'à la valeur générée à l'origine
                        if ParcoursTobEtat <= EnregOrig then
                        begin
                          if TobQBQ.Detail[ParcoursTobEtat].GetValue('QBQ_DATEJOUR') =  DateCalend then
                          begin
                            if Prevu*SumPrevu <> 0
                            then TobQBQ.Detail[ParcoursTobEtat].PutValue('CAPREVU',Prevu*SumPrevu/100)
                            else
                            begin
                              TobQBQ.Detail[ParcoursTobEtat].Free;
                              EnregOrig := EnregOrig-1;
                            end;

                            if ParcoursTobEtat < EnregOrig  then ParcoursTobEtat:=ParcoursTobEtat+1;
                            MoveCurProgressForm('Analyse des calendriers...');
                            FindValue := true
                          end else FindValue := false;
                        end else FindValue := false;

                        //Sinon on ajoute la date dans TobQBQ
                        if ((DateDebPeriod <= TobQBQ.Detail[ParcoursTobEtat].GetValue('QBQ_DATEJOUR')) AND
                        (TobQBQ.Detail[ParcoursTobEtat].GetValue('QBQ_DATEJOUR') <= DateFinPeriod)) then
                        begin
                          if FindValue = false then
                          begin
                            if Prevu <> 0 then
                            begin
                              TobDateSup:= TOB.Create('QBQ',TobQBQ,-1);
                              Try
                                TobDateSup.Dupliquer(TobQBQ.Detail[ParcoursTobEtat],TRUE,TRUE,FALSE);
                                TobDateSup.PutValue('QBQ_DATEJOUR',DateCalend);
                                TobDateSup.PutValue('QBQ_VALAXECT1',ValueAxeCalend[1]);
                                TobDateSup.PutValue('QBQ_VALAXECT2',ValueAxeCalend[2]);
                                TobDateSup.PutValue('QBQ_VALAXECT3',ValueAxeCalend[3]);
                                TobDateSup.PutValue('QBQ_VALAXECT4',ValueAxeCalend[4]);
                                TobDateSup.PutValue('QBQ_VALAXECT5',ValueAxeCalend[5]);
                                TobDateSup.PutValue('QBQ_VALAXECT6',ValueAxeCalend[6]);
                                TobDateSup.PutValue('QBQ_VALAXECT7',ValueAxeCalend[7]);
                                TobDateSup.PutValue('QBQ_VALAXECT8',ValueAxeCalend[8]);
                                TobDateSup.PutValue('QBQ_VALAXECT9',ValueAxeCalend[9]);
                                TobDateSup.PutValue('CAPREVU',Prevu*SumPrevu/100);
                                TobDateSup.PutValue('CAHISTO',0);
                              Finally
                                //TobDateSup.Free;
                              end
                            end
                          end
                        end
                      end
                    end
                  end
                end
              Finally
                TobCurrentCalend.Free
              end
            end
          end
      Finally
        TobCalend.free;
      end;
    end;
  end;
  FiniMoveProgressForm;

  if TraitementOK then
  begin
    if GetCheckBoxState ('DETAILJOUR') = cbChecked
    then TobQBQ.Detail.Sort('QBQ_VALAXECT1;QBQ_VALAXECT2;QBQ_VALAXECT3;QBQ_VALAXECT4;QBQ_VALAXECT5;QBQ_VALAXECT6;QBQ_VALAXECT7;QBQ_VALAXECT8;QBQ_VALAXECT9;QBQ_VALAXECT10;QBQ_DATEJOUR')
    else TobQBQ.Detail.Sort('QBQ_VALAXECT1;QBQ_VALAXECT2;QBQ_VALAXECT3;QBQ_VALAXECT4;QBQ_VALAXECT5;QBQ_VALAXECT6;QBQ_VALAXECT7;QBQ_VALAXECT8;QBQ_VALAXECT9;QBQ_VALAXECT10;QBQ_DATECT');

    {Lancement de l'état}
    CodeEtat := GetControlText( 'FEtat' );

    Case ContextBP of
      0,1,2 : LanceEtatTOB('E','QUF',CodeEtat,TobQBQ,True,False,False,TPageControl(GetControl('Pages')),'',Ecran.Caption,False);
      3 : begin
           InitMoveProgressForm(nil,TraduireMemoire('Edition standard'),TraduireMemoire('Veuillez patienter'), TobQBQ.detail.count,True, False);
           TobQBQ.ParcoursTraitement(['QBQ_CODESESSION'],[EdtSession],true, TobEnLigne);
           if not MoveCurProgressForm(TraduireMemoire('Traitement des données...')) then FiniMoveProgressForm
           else
           begin
             FiniMoveProgressForm;
             TobMultiValQBQ.Detail.Sort('LIBELLE;QBQ_VALAXECT1;QBQ_VALAXECT2;QBQ_VALAXECT3;QBQ_VALAXECT4;QBQ_VALAXECT5;QBQ_VALAXECT6;QBQ_VALAXECT7;QBQ_VALAXECT8;QBQ_VALAXECT9;QBQ_VALAXECT10;QBQ_DATECT');
             LanceEtatTOB('E','QUP',CodeEtat,TobMultiValQBQ,True,False,False,TPageControl(GetControl('Pages')),'',Ecran.Caption,False);
           end;
          end;
    end;
  end;
  InProgress := false;
  Finally
    TobQBQ.Free;
    TobMultiValQBQ.Free;
  end;

end;

{ EVI / Récupération des critères - multiselection}
procedure TOF_QUFQBPBUDGET.DonneFiltre;
var i,Dot,DotTemp:integer;
AxeTemp:hString;
begin
  codeFiltre:='';
  Dot :=0;

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

procedure TOF_QUFQBPBUDGET.RempliTobPourEtatGlobal(const codeSession:string;var laTob:TOB);
var RI,i:integer;
    codeMarche,CodeLibVal : string;
    RS, codeChp, BPInitialise, codeType, stGroupBy, stSQLDate:hString;
    Q:TQuery;
    TabLibelle:array [0..7] of hString;


		{Renseigne une période en lettre}
    function GetPeriodeLibelle:hString;
    begin
      Result:='';
      {quadrimèstre}
      if BPInitialise[1]='7' then
      begin
        result:= ' iif((month(MAX(QBQ_DATECT)) IN (1,2,3,4)), "JANVIER - AVRIL",'
               + ' iif((month(MAX(QBQ_DATECT)) IN (5,6,7,8)), "MAI - AOUT", "SEPTEMBRE - DECEMBRE") ) AS PERIODELIB,';
      end
      {trimèstre}
      else if BPInitialise[1]='6' then
      begin
        result:= ' iif((month(MAX(QBQ_DATECT)) IN (1,2,3)), "JANVIER - MARS",'
               + ' iif((month(MAX(QBQ_DATECT)) IN (4,5,6)), "AVRIL - JUIN",'
               + ' iif((month(MAX(QBQ_DATECT)) IN (7,8,9)), "JUILLET - SEPTEMBRE", "OCTOBRE - DECEMBRE") )) AS PERIODELIB,';
      end;
    end;
    function GetSautsPages:hString;
    begin
      result:= iif(THCheckBox(GetControl('SAUTPAGE1')).Checked=true,'"X" AS SAUTPAGE1,','"-" AS SAUTPAGE1,');
      result:= result+iif(THCheckBox(GetControl('SAUTPAGE2')).Visible=true, iif(THCheckBox(GetControl('SAUTPAGE2')).Checked=true,'"X" AS SAUTPAGE2,','"-" AS SAUTPAGE2,'),'');
      result:= result+iif(THCheckBox(GetControl('SAUTPAGE3')).Visible=true, iif(THCheckBox(GetControl('SAUTPAGE3')).Checked=true,'"X" AS SAUTPAGE3,','"-" AS SAUTPAGE3,'),'');
      result:= result+iif(THCheckBox(GetControl('SAUTPAGE4')).Visible=true, iif(THCheckBox(GetControl('SAUTPAGE4')).Checked=true,'"X" AS SAUTPAGE4,','"-" AS SAUTPAGE4,'),'');
      result:= result+iif(THCheckBox(GetControl('SAUTPAGE5')).Visible=true, iif(THCheckBox(GetControl('SAUTPAGE5')).Checked=true,'"X" AS SAUTPAGE5,','"-" AS SAUTPAGE5,'),'');
      result:= result+iif(THCheckBox(GetControl('SAUTPAGE6')).Visible=true, iif(THCheckBox(GetControl('SAUTPAGE6')).Checked=true,'"X" AS SAUTPAGE6,','"-" AS SAUTPAGE6,'),'');
      result:= result+iif(THCheckBox(GetControl('SAUTPAGE7')).Visible=true, iif(THCheckBox(GetControl('SAUTPAGE7')).Checked=true,'"X" AS SAUTPAGE7,','"-" AS SAUTPAGE7,'),'');
      result:= result+iif(THCheckBox(GetControl('SAUTPAGE8')).Visible=true, iif(THCheckBox(GetControl('SAUTPAGE8')).Checked=true,'"X" AS SAUTPAGE8,','"-" AS SAUTPAGE8,'),'');
      result:= result+iif(THCheckBox(GetControl('SAUTPAGE9')).Visible=true, iif(THCheckBox(GetControl('SAUTPAGE9')).Checked=true,'"X" AS SAUTPAGE9,','"-" AS SAUTPAGE9,'),'');
      result:= result+iif(THCheckBox(GetControl('SAUTPAGE10')).Visible=true,iif(THCheckBox(GetControl('SAUTPAGE10')).Checked=true,'"X" AS SAUTPAGE10,','"-" AS SAUTPAGE10,'),'');
    end;
begin
  RI:=1;
  BPInitialise:='';
  codeMarche := CodeMarcheCommun;
  Q:=OpenSql('SELECT QBS_VUEARBRE,QBS_BPINITIALISE FROM QBPSESSIONBP '+
             'WHERE QBS_CODESESSION="'+codesession+'"',true);
  if not Q.eof then
  begin
    RS:=Q.fields[0].asString;
    BPInitialise:=Q.fields[1].asString;
  end;
  ferme(Q);

  Case ContextBP of
    0,1 : begin //Mode-GC
        	if RS='PTC' then RI:=1
        	else if RS='QTE' then RI:=2
          else if RS='PHT' then RI:=3
          else if RS='UHT' then	RI:=4
          else if RS='UTC' then	RI:=5
        	else if RS='PAH' then	RI:=6
        	else if RS='MAR' then	RI:=7
        end;
    2 : begin
          if RS='DC1' then RI:=1
          else if RS='CD1' then RI:=3
        end;
  end;


  codeChp:='QBQ_CAHISTO,QBQ_CAPREVU,QBQ_CAREALISE ';
  Case ContextBP of
    0,1,2 : begin //Mode-GC-Compta
            //suivant le RI on obtient les champs à récupèrer :
            case RI of
             1 : codeChp:='SUM(QBQ_CAHISTO) AS CAHISTO,SUM(QBQ_CAPREVU) AS CAPREVU,SUM(QBQ_CAREALISE) AS CAREALISE ';
             2 : codeChp:='SUM(QBQ_HISTO) AS CAHISTO,SUM(QBQ_PREVU) AS CAPREVU,SUM(QBQ_REALISE) AS CAREALISE ';
             3 : codeChp:='SUM(QBQ_CAHISTO2) AS CAHISTO,SUM(QBQ_CAPREVU2) AS CAPREVU,SUM(QBQ_CAREALISE2) AS CAREALISE ';
             4 : codeChp:='SUM(QBQ_CAHISTO3) AS CAHISTO,SUM(QBQ_CAPREVU3) AS CAPREVU,SUM(QBQ_CAREALISE3) AS CAREALISE ';
             5 : codeChp:='SUM(QBQ_CAHISTO4) AS CAHISTO,SUM(QBQ_CAPREVU4) AS CAPREVU,SUM(QBQ_CAREALISE4) AS CAREALISE ';
             6 : codeChp:='SUM(QBQ_CAHISTO5) AS CAHISTO,SUM(QBQ_CAPREVU5) AS CAPREVU,SUM(QBQ_CAREALISE5) AS CAREALISE ';
             7 : codeChp:='SUM(QBQ_CAHISTO6) AS CAHISTO,SUM(QBQ_CAPREVU6) AS CAPREVU,SUM(QBQ_CAREALISE6) AS CAREALISE ';
            end;
          end;
      3 : begin //Paie
            codeChp := '';
            LibValAff(codesession,TabLibelle);
            NbValAff := StrToInt(TabLibelle[0]);
            for i := 1 to NbValAff do
            begin
              if i=1 then codeChp:=TabValue[i-1]
              else codeChp:=codeChp + ','+TabValue[i-1];
              if CodeLibVal = '' then CodeLibVal := 'QBS_VALAFFLIB'+IntToStr(i)
                                 else CodeLibVal := CodeLibval + ',QBS_VALAFFLIB'+IntToStr(i);
            end;
            SetLength(TabLibVal,NbValAff+1);
            Q:=OpenSQL('SELECT '+CodeLibVal+' FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+codeSession+'"',true);
            if not Q.eof then
            begin
              for i := 1 to NbValAff do TabLibVal[i]:=Q.Fields[i-1].AsString;
            end;
            Ferme(Q);
          end;
    end; //CASE





  codeType:='';
  if BPInitialise='' then BPInitialise:='0';
  case BPInitialise[1] of
   '2': codeType:='semaine';
   '3': codeType:='quinzaine';
   '4': codeType:='mois';
   '5': codeType:='mois445';
   '6': codeType:='trimestre';
   '7': codeType:='quadrimèstre';
  end;

  {Période d'initialisation de la session}
	LaTob.AddChampSupValeur('SESSION',CodeSession);
	LaTob.AddChampSupValeur('LIBSESSION',RechDom('QUTBPSESSION', CodeSession, False));
	LaTob.AddChampSupValeur('PERIODEINIT',codeType);

  { EVI / Récupération des critères du multiselection }
  DonneFiltre;

  {$IFDEF PAIEGRH}
  ExecuteSQL('DELETE FROM COMMUN WHERE CO_TYPE="QAP" AND CO_CODE NOT IN ("001","010","011")');
  for i := 1 to VH_Paie.PgNbreStatOrg do
  begin
    case i of
     1 : ExecuteSQL('INSERT INTO COMMUN (CO_TYPE,CO_CODE,CO_LIBELLE) VALUES ("QAP","002","'+VH_Paie.PgLibelleOrgStat1+'")');
     2 : ExecuteSQL('INSERT INTO COMMUN (CO_TYPE,CO_CODE,CO_LIBELLE) VALUES ("QAP","003","'+VH_Paie.PgLibelleOrgStat2+'")');
     3 : ExecuteSQL('INSERT INTO COMMUN (CO_TYPE,CO_CODE,CO_LIBELLE) VALUES ("QAP","004","'+VH_Paie.PgLibelleOrgStat3+'")');
     4 : ExecuteSQL('INSERT INTO COMMUN (CO_TYPE,CO_CODE,CO_LIBELLE) VALUES ("QAP","005","'+VH_Paie.PgLibelleOrgStat4+'")');
    end;
  end;
  for i := 1 to VH_Paie.PgNbCombo do
  begin
    case i of
     1 : ExecuteSQL('INSERT INTO COMMUN (CO_TYPE,CO_CODE,CO_LIBELLE) VALUES ("QAP","006","'+VH_Paie.PgLibCombo1+'")');
     2 : ExecuteSQL('INSERT INTO COMMUN (CO_TYPE,CO_CODE,CO_LIBELLE) VALUES ("QAP","007","'+VH_Paie.PgLibCombo2+'")');
     3 : ExecuteSQL('INSERT INTO COMMUN (CO_TYPE,CO_CODE,CO_LIBELLE) VALUES ("QAP","008","'+VH_Paie.PgLibCombo3+'")');
     4 : ExecuteSQL('INSERT INTO COMMUN (CO_TYPE,CO_CODE,CO_LIBELLE) VALUES ("QAP","009","'+VH_Paie.PgLibCombo4+'")');
    end;
  end;
  {$ENDIF PAIEGRH}

  {Remplissage de la Tob}
  { EVI / Ajout de QBQ_UTILISATEUR}
  if GetCheckBoxState ('DETAILJOUR') = cbChecked then
  begin
    codeType := TraduireMemoire ('jour') ;
    stSQLDate := 'QBQ_DATEJOUR,DAY(QBQ_DATEJOUR),';
    if GetControlText( 'FEtat' ) = 'QU0' then
      stGroupBy := ' GROUP BY QBQ_CODESESSION,QBS_LIBSESSION,'+
              'CO1.CO_LIBELLE,CO2.CO_LIBELLE,CO3.CO_LIBELLE,CO4.CO_LIBELLE,CO5.CO_LIBELLE,'+
              'CO6.CO_LIBELLE,CO7.CO_LIBELLE,CO8.CO_LIBELLE,CO9.CO_LIBELLE,CO10.CO_LIBELLE,'+
              'QBQ_VALAXECT1,QBQ_VALAXECT2,QBQ_VALAXECT3,QBQ_VALAXECT4,QBQ_VALAXECT5,QBQ_VALAXECT6,'+
              'QBQ_VALAXECT7,QBQ_VALAXECT8,QBQ_VALAXECT9,QBQ_VALAXECT10, '+
	      'QBQ_LIBVALAXECT1,QBQ_LIBVALAXECT2,QBQ_LIBVALAXECT3,QBQ_LIBVALAXECT4,QBQ_LIBVALAXECT5,QBQ_LIBVALAXECT6,'+
    	      'QBQ_LIBVALAXECT7,QBQ_LIBVALAXECT8,QBQ_LIBVALAXECT9,QBQ_LIBVALAXECT10,QBQ_DATEJOUR,DAY(QBQ_DATEJOUR)' else
      stGroupBy := ' GROUP BY QBQ_CODESESSION,QBS_LIBSESSION,QBQ_DATEJOUR,DAY(QBQ_DATEJOUR),'+
              'CO1.CO_LIBELLE,CO2.CO_LIBELLE,CO3.CO_LIBELLE,CO4.CO_LIBELLE,CO5.CO_LIBELLE,'+
              'CO6.CO_LIBELLE,CO7.CO_LIBELLE,CO8.CO_LIBELLE,CO9.CO_LIBELLE,CO10.CO_LIBELLE,'+
              'QBQ_VALAXECT1,QBQ_VALAXECT2,QBQ_VALAXECT3,QBQ_VALAXECT4,QBQ_VALAXECT5,QBQ_VALAXECT6,'+
              'QBQ_VALAXECT7,QBQ_VALAXECT8,QBQ_VALAXECT9,QBQ_VALAXECT10, '+
	      'QBQ_LIBVALAXECT1,QBQ_LIBVALAXECT2,QBQ_LIBVALAXECT3,QBQ_LIBVALAXECT4,QBQ_LIBVALAXECT5,QBQ_LIBVALAXECT6,'+
    	      'QBQ_LIBVALAXECT7,QBQ_LIBVALAXECT8,QBQ_LIBVALAXECT9,QBQ_LIBVALAXECT10'
  end else
  begin
    stSQLDate := 'QBS_DATEDEBC,QBS_DATEFINC,QBQ_DATECT,';

    stGroupBy := ' GROUP BY QBQ_CODESESSION,QBS_LIBSESSION,QBS_DATEDEBC,QBS_DATEFINC,'+
              'CO1.CO_LIBELLE,CO2.CO_LIBELLE,CO3.CO_LIBELLE,CO4.CO_LIBELLE,CO5.CO_LIBELLE,'+
              'CO6.CO_LIBELLE,CO7.CO_LIBELLE,CO8.CO_LIBELLE,CO9.CO_LIBELLE,CO10.CO_LIBELLE,'+
              'QBQ_VALAXECT1,QBQ_VALAXECT2,QBQ_VALAXECT3,QBQ_VALAXECT4,QBQ_VALAXECT5,QBQ_VALAXECT6,'+
              'QBQ_VALAXECT7,QBQ_VALAXECT8,QBQ_VALAXECT9,QBQ_VALAXECT10, '+
	      'QBQ_LIBVALAXECT1,QBQ_LIBVALAXECT2,QBQ_LIBVALAXECT3,QBQ_LIBVALAXECT4,QBQ_LIBVALAXECT5,QBQ_LIBVALAXECT6,'+
    	      'QBQ_LIBVALAXECT7,QBQ_LIBVALAXECT8,QBQ_LIBVALAXECT9,QBQ_LIBVALAXECT10,QBQ_DATECT';
  end;
  LaTob.LoadDetailDBFromSQL('QBPCUBETMP',
  	      'SELECT QBQ_CODESESSION,QBS_LIBSESSION,' + stSQLDate +
              'CO1.CO_LIBELLE AS LIBAXE1,CO2.CO_LIBELLE AS LIBAXE2,CO3.CO_LIBELLE AS LIBAXE3,'+
              'CO4.CO_LIBELLE AS LIBAXE4,CO5.CO_LIBELLE AS LIBAXE5,CO6.CO_LIBELLE AS LIBAXE6,'+
              'CO7.CO_LIBELLE AS LIBAXE7,CO8.CO_LIBELLE AS LIBAXE8,CO9.CO_LIBELLE AS LIBAXE9,CO10.CO_LIBELLE AS LIBAXE10,'+
              'QBQ_VALAXECT1,QBQ_VALAXECT2,QBQ_VALAXECT3,QBQ_VALAXECT4,QBQ_VALAXECT5,QBQ_VALAXECT6,'+
        	    'QBQ_VALAXECT7,QBQ_VALAXECT8,QBQ_VALAXECT9,QBQ_VALAXECT10, '+
          	  'QBQ_LIBVALAXECT1,QBQ_LIBVALAXECT2,QBQ_LIBVALAXECT3,QBQ_LIBVALAXECT4,QBQ_LIBVALAXECT5,'+
            	'QBQ_LIBVALAXECT6,QBQ_LIBVALAXECT7,QBQ_LIBVALAXECT8,QBQ_LIBVALAXECT9,QBQ_LIBVALAXECT10,'+
 	            '"'+CodeType+'" AS PERIODE,1 AS GR11,'+
               GetPeriodeLibelle+GetSautsPages+codeChp+' FROM QBPCUBETMP '+
              'LEFT JOIN QBPSESSIONBP ON (QBS_CODESESSION=QBQ_CODESESSION) ' +
              'LEFT JOIN COMMUN CO1 ON (CO1.CO_CODE=QBS_CODEAXES1 AND CO1.CO_TYPE="'+codeMarche+'") ' +
              'LEFT JOIN COMMUN CO2 ON (CO2.CO_CODE=QBS_CODEAXES2 AND CO2.CO_TYPE="'+codeMarche+'") ' +
              'LEFT JOIN COMMUN CO3 ON (CO3.CO_CODE=QBS_CODEAXES3 AND CO3.CO_TYPE="'+codeMarche+'") ' +
              'LEFT JOIN COMMUN CO4 ON (CO4.CO_CODE=QBS_CODEAXES4 AND CO4.CO_TYPE="'+codeMarche+'") ' +
              'LEFT JOIN COMMUN CO5 ON (CO5.CO_CODE=QBS_CODEAXES5 AND CO5.CO_TYPE="'+codeMarche+'") ' +
              'LEFT JOIN COMMUN CO6 ON (CO6.CO_CODE=QBS_CODEAXES6 AND CO6.CO_TYPE="'+codeMarche+'") ' +
              'LEFT JOIN COMMUN CO7 ON (CO7.CO_CODE=QBS_CODEAXES7 AND CO7.CO_TYPE="'+codeMarche+'") ' +
              'LEFT JOIN COMMUN CO8 ON (CO8.CO_CODE=QBS_CODEAXES8 AND CO8.CO_TYPE="'+codeMarche+'") ' +
              'LEFT JOIN COMMUN CO9 ON (CO9.CO_CODE=QBS_CODEAXES9 AND CO9.CO_TYPE="'+codeMarche+'") ' +
              'LEFT JOIN COMMUN CO10 ON (CO10.CO_CODE=QBS_CODEAXES10 AND CO10.CO_TYPE="'+codeMarche+'") ' +
      	      'WHERE QBQ_CODESESSION="'+codeSession+'" '+
              ' AND QBQ_UTILISATEUR="'+V_PGI.User +'" '+codefiltre+ stGroupBy) ;

  {$IFDEF PAIEGRH}
  ExecuteSQL('DELETE FROM COMMUN WHERE CO_TYPE="QAP" AND CO_CODE NOT IN ("001","010","011")');
  {$ENDIF PAIEGRH}
end;

procedure TOF_QUFQBPBUDGET.FEtatOnClick(Sender: TObject);
begin
  if ContextBP <> 3 then
  begin
    if TFQRS1( Ecran ).FEtat.Value = 'QU0' then
    begin
      SetControlVisible( 'VOIRCUMULS', True );
      SetControlEnabled( 'DETAILJOUR', True );
    end else
    if TFQRS1( Ecran ).FEtat.Value = 'QUJ' then
    begin
      SetControlVisible( 'VOIRCUMULS', False );
      SetControlChecked( 'DETAILJOUR', True );
      SetControlEnabled( 'DETAILJOUR', False );
    end;
  end;
end;

procedure TOF_QUFQBPBUDGET.FEtatOnChange( Sender : TObject );
begin
  FEtatOnClick (Ecran);
end;

procedure TOF_QUFQBPBUDGET.OnNouvelleRecherche(Sender: TObject);
begin
  VideFiltre(TFQRS1 (Ecran).FFiltres, TFQRS1 (Ecran).Pages);
  TFQRS1 (Ecran).ListeFiltre.New;
  ComboModeleEtat.ItemIndex := 0;
  FEtatOnClick (Ecran);
end;

procedure TOF_QUFQBPBUDGET.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = VK_ESCAPE) and (InProgress=true) then Key := 0;
  if( (Key = VK_F9) and Assigned(BValider) ) then BValider.Click;
end;

Initialization
  registerclasses ( [ TOF_QUFQBPBUDGET ] ) ;
end.

