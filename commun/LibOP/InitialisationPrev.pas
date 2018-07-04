unit InitialisationPrev;

interface
uses SysUtils,Controls,HMsgBox,HEnt1,HCtrls,
     BPInitArbre,BPBasic,BPFctSession;

const
  TexteMessage : array[1..26] of hstring =  ( {MODIF MODE UNICODE : hstring au lieu de string}
    {1}'Pour le calcul d''initialisation par jour : %s ',
    {2}'La période de référence et la période d''objectif doivent compter le même nombre de jours.',

    {3}'Pour le calcul d''initialisation par semaine : %s ',
    {4}'Les dates de début doivent correspondre au premier jour de la semaine définie en paramètre (%0:s). %1:s ', {MODIF MODE UNICODE}
    {5}'Les dates de fin doivent correspondre au dernier jour de la semaine précédente (%s).',
    {6}'Le nombre de semaines pour la période d''objectif (%0:s) doit être égal à celui de la période de référence (%1:s).', {MODIF MODE UNICODE}

    {7}'Pour le calcul d''initialisation par quinzaine : %s ',
    {8}'Les dates de début doivent correspondre au premier jour ou au 16 du mois de début. %s ',
    {9}'Les dates de fin doivent correspondre au 15 ou au dernier jour du mois de fin. ',
    {10}'Le nombre de quinzaines pour la période d''objectif (%0:s) doit être égal à celui de la période de référence (%1:s).', {MODIF MODE UNICODE}

    {11}'Pour le calcul d''initialisation par mois : %s ',
    {12}'Les dates de début courante et de réference doivent correspondre au premier jour du mois. %s ',
    {13}'Les dates de fin courante et de réference doivent correspondre au dernier jour du mois. ',
    {14}'Le nombre de mois pour la période d''objectif (%0:s) doit être égal à celui de la période de référence (%1:s).', {MODIF MODE UNICODE}

    {15}'Pour le calcul d''initialisation par mois 4-4-5 : %s ',
    {16}'Les dates de début courante et de réference doivent correspondre au premier jour d''une période.%s ',
    {17}'Les dates de fin courante et de réference doivent correspondre au dernier jour d''une période.',
    {18}'Le nombre de mois 4-4-5 pour la période d''objectif (%0:s) doit être égal à celui de la période de référence (%1:s).', {MODIF MODE UNICODE}

    {19}'Pour le calcul d''initialisation par trimestre : %s ',
    {20}'Les dates de début doivent correspondre au premier jour d''un trimestre (%0:s,%1:s,%2:s,%3:s).%4:s ', {MODIF MODE UNICODE}
    {21}'Les dates de fin doivent correspondre au dernier jour d''un trimestre (%0:s,%1:s,%2:s,%3:s).', {MODIF MODE UNICODE}
    {22}'Le nombre de trimestres pour la période d''objectif (%0:s) doit être égal à celui de la période de référence (%1:s).', {MODIF MODE UNICODE}

    {23}'Pour le calcul d''initialisation par quadrimestre : %s ',
    {24}'Les dates de début doivent correspondre au premier jour d''un quadrimestre (%0:s,%1:s,%2:s).%3:s ', {MODIF MODE UNICODE}
    {25}'Les dates de fin doivent correspondre au dernier jour d''un quadrimestre (%0:s,%1:s,%2:s).', {MODIF MODE UNICODE}
    {26}'Le nombre de quadrimestres pour la période d''objectif (%0:s) doit être égal à celui de la période de référence (%1:s).' {MODIF MODE UNICODE}
    );

procedure InitialiseGlobal(const codeSession:hString);
function InitialiseDelai(const codeSession,codeMaille,SaiRef,SaiC:hString;
          DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime):boolean;

procedure IntialiseCoeff(const codeSession,BPInitialise:hString;coeff1:boolean);

implementation

procedure IntialiseCoeff(const codeSession,BPInitialise:hString;coeff1:boolean);
begin
 InitArbre(codeSession,BPInitialise,coeff1,false);
 ModifSession(codeSession,'X');
end;


procedure InitialiseGlobal(const codeSession:hString);
begin
 InitArbre(codeSession,'0',false,false);
end;

function InitialiseDelai(const codeSession,codeMaille,SaiRef,SaiC:hString;
          DateDebC,DateFinC,DateDebRef,DateFinRef:TDateTime):boolean;
var rep,PremierJourS,JourF,NbMailleRef,NbMailleC:integer;
    Maille,JourS,JourFS:hString;
    nbJourR,nbJourC:double;
    PerC,PerR,PerCF,PerRF:TPeriodeP;
    OkQuinz,okTrim,okQuadri:boolean;
    an,mois,day:word;
    okQuest:boolean;
begin
 rep:=mrNo;
 result := false;
 okQuest:=true;
 //si maille = jour
 if codeMaille='1'
  then
   begin
    nbJourR:=DateFinRef-DateDebRef;
    nbJourC:=DateFinC-DateDebC;
    if (nbJourR<>nbJourC)
     then PGIINFO( Format_(TraduireMemoire( TexteMessage[1] ),['#13#10']) + TraduireMemoire(TexteMessage[2])) {MODIF MODE UNICODE : Ajout TraduireMemoire et format_}
     else rep:=HShowmessage('1;Initialisation;Vous allez initialiser la session par jour.#13#10 Etes-vous sûr ?;Q;YN;N;N', '', '')
   end;
 //si maille = semaine
 if codeMaille='2'
  then
   begin
    PremierJourS:=V_PGI.PremierJourSemaine;
    case PremierJourS of
     1 : begin
          JourS:=TraduireMemoire('Dimanche');
          JourFS:=TraduireMemoire('Samedi');
          JourF:=7;
         end;
     2 : begin
          JourS:=TraduireMemoire('Lundi');
          JourFS:=TraduireMemoire('Dimanche');
          JourF:=1;
         end;
     7 : begin
          JourS:=TraduireMemoire('Samedi');
          JourFS:=TraduireMemoire('Vendredi');
          JourF:=6;
         end;
     else
      begin
       JourS:=TraduireMemoire('Lundi');
       JourF:=1;
      end;
    end;
    if (SaiRef='') and ((DayOfWeek(DateDebC)<>PremierJourS) or (DayOfWeek(DateFinC)<>JourF)
       or (DayOfWeek(DateDebRef)<>PremierJourS) or (DayOfWeek(DateFinRef)<>JourF))
     then PGIINFO(Format_(TraduireMemoire(TexteMessage[3]),[#13#10]) + Format_(TraduireMemoire(TexteMessage[4]),[JourS,#13#10]) + Format_(TraduireMemoire(TexteMessage[5]),[JourFS])) {MODIF MODE UNICODE : Ajout TraduireMemoire et format_}
     else if (DateDebC > DateFinC) or (DateDebRef > DateFinRef)  { EVI / Contrôle supplémentaire }
     then
       begin
         PGIINFO(TraduireMemoire('Les dates sont incohérentes.'),''); {MODIF MODE UNICODE : Ajout TraduireMemoire}
       end
     else
       begin
        //test nombre de mailles égales
        NbMailleC:=NbSemaineIntervalle(DateDebC,DateFinC);
        NbMailleRef:=NbSemaineIntervalle(DateDebRef,DateFinRef);
        if (SaiRef='') and (NbMailleC<>NbMailleRef)
        then
        begin
          PGIINFO(Format_(TraduireMemoire(TexteMessage[3]),[#13#10]) + Format_(TraduireMemoire(TexteMessage[6]),[IntToStr(NbMailleC),IntTostr(NbMailleRef)])); {MODIF MODE UNICODE : Ajout TraduireMemoire et format_}
          okQuest:=false;
        end;
        if OkQuest
        then rep:=HShowmessage('1;Initialisation;Vous allez initialiser la session par semaine.#13#10 Etes-vous sûr ?;Q;YN;N;N', '', '')
       end;
   end;

 //si maille = qunizaine
 if (codeMaille='3')
  then
   begin
    OkQuinz:=true;
    DecodeDate(DateDebC,an,mois,day);
    if (day<>16) and (day<>1)
     then OkQuinz:=false;
    DecodeDate(DateDebRef,an,mois,day);
    if (day<>16) and (day<>1)
     then OkQuinz:=false;
    DecodeDate(DateFinC,an,mois,day);
    if (day<>15) and (DateFinC<>(FINDEMOIS(DateFinC)))
     then OkQuinz:=false;
    DecodeDate(DateFinRef,an,mois,day);
    if (day<>15) and (DateFinRef<>(FINDEMOIS(DateFinRef)))
     then OkQuinz:=false;

    if (SaiRef<>'') and (SaiC<>'')
     then OkQuinz:=true;

    if OkQuinz=false
     then PGIINFO(Format_(TraduireMemoire(TexteMessage[7]),[#13#10]) + Format_(TraduireMemoire(TexteMessage[8]),[#13#10]) + TraduireMemoire(TexteMessage[9])) {MODIF MODE UNICODE : Ajout TraduireMemoire et format_}
     else
       begin
        //test nombre de mailles égales
        NbMailleC:=NbQuinzaineIntervalle(DateDebC,DateFinC);
        NbMailleRef:=NbQuinzaineIntervalle(DateDebRef,DateFinRef);
        if (SaiRef='') and (NbMailleC<>NbMailleRef)
        then
        begin
          PGIINFO(Format_(TraduireMemoire(TexteMessage[7]),[#13#10]) + Format_(TraduireMemoire(TexteMessage[10]),[IntToStr(NbMailleC),IntTostr(NbMailleRef)])); {MODIF MODE UNICODE : Ajout TraduireMemoire et format_}
          okQuest:=false;
        end;
        if OkQuest
        then rep:=HShowmessage('1;Initialisation;Vous allez initialiser la session par quinzaine.#13#10 Etes-vous sûr ?;Q;YN;N;N', '', '')
      end;
   end;

 //si maille = mois
 if (codeMaille='4')
  then
   begin
     if (SaiRef='') and ((DateDebC<>(DEBUTDEMOIS(DateDebC))) or (DateDebRef<>(DEBUTDEMOIS(DateDebRef)))
     or (DateFinC<>(FINDEMOIS(DateFinC))) or (DateFinRef<>(FINDEMOIS(DateFinRef))))
     then PGIINFO(Format_(TraduireMemoire(TexteMessage[11]),[#13#10]) + Format_(TraduireMemoire(TexteMessage[12]),[#13#10]) + TraduireMemoire(TexteMessage[13])) {MODIF MODE UNICODE : Ajout TraduireMemoire et format_}
     else
       begin
        //test nombre de mailles égales
        NbMailleC:=NbMoisIntervalle(DateDebC,DateFinC);
        NbMailleRef:=NbMoisIntervalle(DateDebRef,DateFinRef);
        if (SaiRef='') and (NbMailleC<>NbMailleRef)
        then
        begin
          PGIINFO(Format_(TraduireMemoire(TexteMessage[11]),[#13#10]) + Format_(TraduireMemoire(TexteMessage[14]),[IntToStr(NbMailleC),IntTostr(NbMailleRef)])); {MODIF MODE UNICODE : Ajout TraduireMemoire et format_}
          okQuest:=false;
        end;
        if okQuest
        then rep:=HShowmessage('1;Initialisation;Vous allez initialiser la session par mois.#13#10 Etes-vous sûr ?;Q;YN;N;N', '', '')
      end;
   end;
 //si maille = mois 4-4-5
 if (codeMaille='5')
  then
   begin
    PerC:=FindPeriodeP(DateDebC);
    PerR:=FindPeriodeP(DateDebRef);
    PerCF:=FindPeriodeP(DateFinC-1);
    PerRF:=FindPeriodeP(DateFinRef-1);

    if (SaiRef='') and ((DateDebC<>(PerC.datedeb)) or (DateDebRef<>(PerR.datedeb))
        or (DateFinC<>(PerCF.datefin)) or (DateFinRef<>(PerRF.datefin)))
     then PGIINFO(Format_(TraduireMemoire(TexteMessage[15]),[#13#10]) + Format_(TraduireMemoire(TexteMessage[16]),[#13#10]) + TraduireMemoire(TexteMessage[17])) {MODIF MODE UNICODE : Ajout TraduireMemoire et format_}
     else
      begin
        //test nombre de mailles égales
        NbMailleC:=DonneNbIntervalleMailleDateDebDateFin('5',DateDebC,DateFinC);
        NbMailleRef:=DonneNbIntervalleMailleDateDebDateFin('5',DateDebRef,DateFinRef);
        if (SaiRef='') and (NbMailleC<>NbMailleRef)
        then
        begin
          PGIINFO(Format_(TraduireMemoire(TexteMessage[15]),[#13#10]) + Format_(TraduireMemoire(TexteMessage[18]),[IntToStr(NbMailleC),IntTostr(NbMailleRef)])); {MODIF MODE UNICODE : Ajout TraduireMemoire et format_}
          okQuest:=false;
        end;
        if okQuest
        then rep:=HShowmessage('1;Initialisation;Vous allez initialiser la session par mois 4-4-5.#13#10 Etes-vous sûr ?;Q;YN;N;N', '', '')
      end;
   end;

 //si maille = trimestre
 if (codeMaille='6')
  then
   begin
    OkTrim:=true;
    DecodeDate(DateDebC,an,mois,day);
    if (day<>1) and ((mois<>1) or (mois<>4) or (mois<>7) or (mois<>10))
     then OkTrim:=false;
    DecodeDate(DateDebRef,an,mois,day);
    if (day<>1) and ((mois<>1) or (mois<>4) or (mois<>7) or (mois<>10))
     then OkTrim:=false;
    DecodeDate(DateFinC,an,mois,day);
    if not( ((mois=3) and (day=31)) or ((mois=6) and (day=30)) or
            ((mois=9) and (day=30)) or ((mois=12) and (day=31)) )
     then OkTrim:=false;
    DecodeDate(DateFinRef,an,mois,day);
    if not( ((mois=3) and (day=31)) or ((mois=6) and (day=30)) or
            ((mois=9) and (day=30)) or ((mois=12) and (day=31)) )
     then OkTrim:=false;

    if (SaiRef<>'') and (SaiC<>'')
     then okTrim:=true;

    if okTrim=false
     then PGIINFO(Format_(TraduireMemoire(TexteMessage[19]),[#13#10]) +
                 Format_(TraduireMemoire(TexteMessage[20]),[TraduitDateFormat('01/01'),TraduitDateFormat('01/04'),TraduitDateFormat('01/07'),TraduitDateFormat('01/10'),#13#10]) +
                 Format_(TraduireMemoire(TexteMessage[21]),[TraduitDateFormat('31/03'),TraduitDateFormat('30/06'),TraduitDateFormat('30/09'),TraduitDateFormat('31/12')])) {MODIF MODE UNICODE : Ajout TraduireMemoire et format_}
     else
       begin
        //test nombre de mailles égales
        NbMailleC:=DonneNbIntervalleMailleDateDebDateFin('6',DateDebC,DateFinC);
        NbMailleRef:=DonneNbIntervalleMailleDateDebDateFin('6',DateDebRef,DateFinRef);
        if (SaiRef='') and (NbMailleC<>NbMailleRef)
        then
        begin
          PGIINFO(Format_(TraduireMemoire(TexteMessage[19]),[#13#10]) + Format_(TraduireMemoire(TexteMessage[22]),[IntToStr(NbMailleC),IntTostr(NbMailleRef)])); {MODIF MODE UNICODE : Ajout TraduireMemoire et format_}
          okQuest:=false;
        end;
        if okQuest
        then rep:=HShowmessage('1;Initialisation;Vous allez initialiser la session par trimestre.#13#10 Etes-vous sûr ?;Q;YN;N;N', '', '')
       end;
   end;

 //si maille = quadrimestre
 if (codeMaille='7')
  then
   begin
    OkQuadri:=true;
    DecodeDate(DateDebC,an,mois,day);
    if (day<>1) and ((mois<>1) or (mois<>5) or (mois<>9))
     then OkQuadri:=false;
    DecodeDate(DateDebRef,an,mois,day);
    if (day<>1) and ((mois<>1) or (mois<>5) or (mois<>9))
     then OkQuadri:=false;
    DecodeDate(DateFinC,an,mois,day);
    if not( ((mois=4) and (day=30)) or ((mois=8) and (day=31)) or
            ((mois=12) and (day=31)) )
     then OkQuadri:=false;
    DecodeDate(DateFinRef,an,mois,day);
    if not( ((mois=4) and (day=30)) or ((mois=8) and (day=31)) or
            ((mois=12) and (day=31)) )
     then OkQuadri:=false;

    if (SaiRef<>'') and (SaiC<>'')
     then okQuadri:=true;

    if okQuadri=false
     then PGIINFO(Format_(TraduireMemoire(TexteMessage[23]), [#13#10]) +
                  Format_(TraduireMemoire(TexteMessage[24]), [TraduitDateFormat('01/01'),TraduitDateFormat('01/05'),TraduitDateFormat('01/09'),#13#10]) +
                  Format_(TraduireMemoire(TexteMessage[25]), [TraduitDateFormat('30/04'),TraduitDateFormat('31/08'),TraduitDateFormat('31/12')])) {MODIF MODE UNICODE : Ajout TraduireMemoire et format_}
     else
      begin
        //test nombre de mailles égales
        NbMailleC:=DonneNbIntervalleMailleDateDebDateFin('7',DateDebC,DateFinC);
        NbMailleRef:=DonneNbIntervalleMailleDateDebDateFin('7',DateDebRef,DateFinRef);
        if (SaiRef='') and (NbMailleC<>NbMailleRef)
        then
        begin
          PGIINFO(Format_(TraduireMemoire(TexteMessage[23]),[#13#10]) + Format_(TraduireMemoire(TexteMessage[26]),[IntToStr(NbMailleC),IntTostr(NbMailleRef)])); {MODIF MODE UNICODE : Ajout TraduireMemoire et format_}
          okQuest:=false;
        end;
        if okQuest
        then rep:=HShowmessage('1;Initialisation;Vous allez initialiser la session par quadrimestre.#13#10 Etes-vous sûr ?;Q;YN;N;N', '', '')
      end;
   end;

 if rep=MrYes
  then
   begin
    if QuestBPInit(codeSession)
     then
       if QuestBPCalend(codeSession)
        then
        begin
         InitArbre(codeSession,codeMaille,false,false);
         result := true;
        end;
   end;

  //si maille = MAJ
  if (codeMaille='X')
  then
  begin
    Maille:=SessionBPInitialise(codeSession);
    if Maille<>'' then
    begin
      InitArbre(codeSession,Maille,false,true);
      HShowmessage('1;Session Objectif;Mise à jour effectuée.;X;O;O;O;', '', '');
    end;
  end;
end;


end.
