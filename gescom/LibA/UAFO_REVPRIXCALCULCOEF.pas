unit UAFO_REVPRIXCALCULCOEF;

interface

  uses

  stdCtrls, Controls, Classes,  forms, sysutils, ComCtrls,
  HCtrls, HEnt1, HMsgBox, UTOF, DicoBTP, UTob, HTB97,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  paramsoc,
  formule,
  //utilplanning,
  utilRevision,
  RT_Parser,
  UTilFonctionCalcul;

Type
    TDerniereOuProchaine =(Non,Derniere,Prochaine);

    TCALCULCOEF = class
    private

      //formule
      fStFormule      : String;
      fStCodeFormule  : String;
      fStLibFormule   : String;
      fStTypeFormule  : String;
      fArCodeIndice   : array [1 .. 10] of String;
      fInArrondiCoeff : Integer;
      fInArrondiAff   : Integer;
      fStTypeIndice   : String;
      fInNbMois       : Integer;
      fBoFormuleOk    : Boolean;

      //paramFormule
      fStAffaire          : String;

      fArValInitIndice    : array [1 .. 10] of Double;
      fArCodeIndiceAff    : array [1 .. 10] of String;
      fArPubCode          : array [1 .. 10] of String;
      fDtLastDateRevApp   : TDateTime;
      fDtLastDateRevCalc  : TDateTime;
      fDtNextDateRevCalc  : TDateTime;
      fDtNextDateApp      : TDateTime;
      fDtInitiale         : TDateTime;
      fDtPremiereRev      : TDateTime;
      fBoAppliqCoefAuto   : Boolean;
      fStPeriodeRev       : String;
      fStModeLecture      : String;
      fInMoisLecture      : Integer;
      fBoFinMois          : Boolean;
      fInJourLecture      : Integer;
      fRdSeuilMini        : Double;
      fRdSeuilMaxi        : Double;
      fBoParamFormuleOK   : Boolean;

      // val indice
      fArValIndice        : array [1 .. 10] of Double;
      fArDatIndice        : array [1 .. 10] of TDateTime;
      fArIndPubCode       : Array [1 .. 10] of String;
      fArIndDateVal       : Array [1 .. 10] of TDateTime;

      fRdCoef             : Double;
      fInNbindices        : Integer;
      fBoCoefRegul        : Boolean;
      fBoProchaine        : Boolean;
      fDtEcheance         : TDateTime; // facturation
      fDtLectureIndice    : TDateTime; // date de lecture des indices
      fSLLog              : TStringList;
      fDtDateRevision     : TDateTime;
      //fBoAncienCoef       : Boolean;

      function LoadValeurInitIndices(var pStIndiceErreur : String) : Boolean; // chargement des valeurs initiales des indices
      Function LoadValeursIndices(var pStIndiceErreur : String) : Boolean;
                                                        
      Function ValeurIndice(pDate : TDateTime; pStIndice, pStPubCode : String;
                            pBoLastIndice, pBoDeuxDates, pBoInitial : Boolean;
                            var pDtIndice, pDtIndDateVal : TDateTime;
                            var pStIndPubCode : String;  var pStIndiceErreur : String) : Double; // retourne la valeur indice dans la base d'indice

      Function ValeursIndiceUtilise(pInNbValeurs : Integer; pStIndice : String; var pArValeurs : Array of double) : Integer; // retourne le nombre de valeurs d'indices trouvées
      Function ValeursIndiceSaisis(pInNbValeurs : Integer; pStIndice, pStPubCode : String; var pArValeurs : Array of double) : Integer;  // retourne le nombre de valeurs d'indices trouvées
      function RangIndice(pStIndice : String) : Integer;
      procedure CalculNbIndices;

      function DateLectureIndices : TDateTime; //recherche des dates de lecture des indices
      function CalculCoef : boolean; // calcul du coefficient
      function TestFormuleActualisation : boolean;
      function testActualisation : boolean;
      //procedure DecodeTermes(pStFormule : String; var vArTermes, vArArrondis : Array of String);
      //function SetValIndiceFormule(pStIndice : String) : variant;
      //function IndiceZero(pStIndice : String) : Boolean;

      Function InsertCoefficients(pBoRecupTob : Boolean; pTobParamFormule, pTobRevision : Tob) : Integer;
      Function InsertCoefRegul(pDateRegul : TDateTime) : Integer;      
      function MaxRevision : Integer;
      procedure AddLigneCoef(pTob : Tob; pInNumFormule : Integer; var pInCpt : Integer);
      procedure OnGetVar( Sender: TObject; VarName: String; VarIndx: Integer;var Value: variant) ;
      //procedure OnSetVar( Sender: TObject; VarName: String; VarIndx: Integer;Value: variant) ;
      //procedure OnGetList( Sender: TObject; IdentList: String;var List: TStringList );
      procedure OnFunction( Sender: TObject; FuncName: String; Params: Array of variant ; var Result: Variant );
      //procedure OnInitialization( Sender: TObject );
      //procedure OnFinalization( Sender: TObject );
      
    public
      function CalculCoefFacturation(pStAffaire : String; pTobFormule, pTobParamFormule, pTobRevision : Tob; pBoRecupTob : Boolean): Boolean;
      function CalculCoefRegularisation(pStAffaire : String; pDtEcheance : TDateTime): Boolean;
      function CalculCoefFormule(pStAffaire, pStCodeFormule : String;
                                 pDtEcheance : TDateTime; pBoLog, pBoMsg : Boolean;
                                 pBoCoefRegul, pBoProchaine : Boolean;
                                 pTobParamFormule, pTobRevision : Tob; pBoRecupTob : Boolean): Boolean;
      procedure GestionErreur(pBoLog, pBoMsg : Boolean; pTexteMsg : String);
      function LoadFormule(pStAffaire, pStCodeFormule : String): Boolean; // chargement données formule
      function MajDatesRevisionsApresApplication(pBoForcerApplication, pBoRecupTob : boolean; pTobParamFormule : Tob) : Boolean; // mise a jour des dates de revision apres une revision
      procedure AppliquerCoef(pBoAncienCoef, pBoMajDates : Boolean);
      function DesappliquerCoef(pDtDate : TDateTime) : Boolean;
      function LectureValeursInit(pStAffaire, pStCodeFormule : String;
                                  var pArValInitIndice : Array of double): Boolean;
      function CalculFormuleAuto(pTob : Tob) : Boolean;
      function DerniereOuProchaine(pStAffaire, pStFormule : String) : TDerniereOuProchaine;
      function FormuleEditionDetail(pStAffaire, pStCodeFormule : String; var pStFormule : String) : boolean;
      
   end;

const
	TexteMsg : array[1..16] of string 	= (
          {1}        'Le coefficient ne peut pas être calculé, il manque des valeurs initiales d''indice.',
          {2}        'Erreur lors de l''enregistrement du coefficient.',
          {3}        'Affaire : %s. La formule %s : "%s" n''est pas valide.',
          {4}        'Affaire : %s. Une erreur s''est produite lors du chargement des données de la formule.',
          {5}        'Affaire : %s. Le chargement des valeurs initiales de l''indice %s a été interrompu.' + #13#10 + 'Vérifier que vous avez bien saisi toutes les valeurs d''indices nécessaires au calcul de la formule.',
          {6}        'Affaire : %s. Le chargement des valeurs de l''indice %s a été interrompu.' + #13#10 + 'Vérifier que vous avez bien saisi toutes les valeurs d''indices nécessaires au calcul de la formule.',
          {7}        'Affaire : %s. Le coefficient ne peut pas être calculé car la formule %s : "%s" n''est pas correcte.',
          {8}        'Affaire : %s. Le paramétrage de la formule %s : "%s" n''a pas été effectué.',
          {9}        'Affaire : %s. Calcul des coefficients de la formule %s : "%s".',
          {10}       'Affaire : %s. Recalcul des coefficients de la formule %s : "%s".',
          {11}       'Calcul d''une actualisation.',
          {12}       'Affaire : %s. La formule d''actualisation n''a pas a être calculée.'+ #13#10 + ' La Date de d''acceptation n''est pas postérieure de 3 mois à la date de début de l''affaire',
          {13}       'Affaire : %s. Impossible de recalculer la formule %s : "%s". ' + #13#10 + 'Des coefficients de la formule sont déjà appliqués.',
          {14}       'Erreur lors du calcul du coefficient',
          {15}       'Affaire : %s. On ne peut pas annuler l''application du coefficient de la formule %s à la date du %s, car des documents existent pour ce coefficient.',
          {16}       'Affaire : %s. On ne peut pas annuler l''application du coefficient de la formule %s à la date du %s, car des coefficients ont été calculés à une date ultérieure.'
          );

  cStIndiceZero   = '°';
  cStOuvreArrondi = 'ARR{';
  cStSepArrondi   = ';';
  cStFermeArrondi = '}';
  cStVide         = '...';
  cInVide         = '0';
  cInFormule1     = 1;
  cInFormule2     = 2;

  cBoDeuxDates      = True;
  cBoDtApplication  = False;
  cBoDernierIndice  = True;
  cBoProchainIndice = False;

implementation

//uses Windows,ExtCtrls,RichEdit;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 03/04/2003
Modifié le ... :
Description .. : Chargement des données d'une formule pour une affaire donnée
Mots clefs ... :
*****************************************************************}
function TCALCULCOEF.LoadFormule(pStAffaire, pStCodeFormule : String) : Boolean;
var
  vQr : TQuery;
  vSt : String;
  i   : Integer;

begin

  result := true;
  vSt := 'SELECT * FROM AFFORMULE, AFPARAMFORMULE ';
  vSt := vSt + 'WHERE AFC_AFFAIRE = "' + pStAffaire + '" ';
  vSt := vSt + 'AND AFC_FORCODE = "' + pStCodeFormule + '" ';
  vSt := vSt + 'AND AFC_FORCODE = AFE_FORCODE ';

  vQr := nil;
  Try
    Try
      vQR := OpenSql(vSt, True,-1,'',true);
      if Not vQR.Eof then
      begin
        if vQr.FindField('AFE_FORMULEOK').AsString = 'X' then
        begin
          fStCodeFormule  := vQr.FindField('AFE_FORCODE').AsString;
          fStLibFormule   := vQr.FindField('AFE_FORLIB').AsString;
          fStFormule := vQr.FindField('AFE_FOREXPRESSION').AsString;
          if fStFormule = '' then
            result := False
          else
          begin

            fStTypeFormule  := vQr.FindField('AFE_FORTYPE').AsString;
            for i := 1 to 10 do
              fArCodeIndice[i] := vQr.FindField('AFE_INDCODE' + intToSTr(i)).AsString;

            CalculNbIndices;

            fInArrondiCoeff := vQr.FindField('AFE_ARRONDICOEF').AsInteger;
            fInArrondiAff   := vQr.FindField('AFE_NBDECCOEF').AsInteger;
            fStTypeIndice   := vQr.FindField('AFE_INDTYPE').AsString;
            fInNbMois       := vQr.FindField('AFE_NBMOIS').AsInteger;
            fBoFormuleOk    := vQr.FindField('AFE_FORMULEOK').AsString = 'X';

            //paramFormule
            for i := 1 to fInNbIndices do
              fArCodeIndiceAff[i] := vQr.FindField('AFC_INDAFF' + intToSTr(i)).AsString;

            for i := 1 to fInNbIndices do
              fArPubCode[i] := vQr.FindField('AFC_PUBCODE' + intToSTr(i)).AsString;

            for i := 1 to fInNbIndices do
              fArValInitIndice[i] := vQr.FindField('AFC_VALINITIND' + intToSTr(i)).AsFloat;

            fStAffaire          := vQr.FindField('AFC_AFFAIRE').AsString;
            fDtLastDateRevApp   := vQr.FindField('AFC_LASTDATEAPP').AsDateTime;
            fDtLastDateRevCalc  := vQr.FindField('AFC_LASTDATECALC').AsDateTime;
            fDtNextDateRevCalc  := vQr.FindField('AFC_NEXTDATECALC').AsDateTime;
            fDtNextDateApp      := vQr.FindField('AFC_NEXTDATEAPP').AsDateTime;

            //
            if vQR.FindField('AFC_NEXTDATEAPP').AsDateTime = iDate2099 then
              fDtNextDateApp := vQR.FindField('AFC_PREMIEREDATE').AsDateTime
            else
              fDtNextDateApp := vQR.FindField('AFC_NEXTDATEAPP').AsDateTime;

            if fBoProchaine then
              fDtDateRevision := fDtNextDateApp
            else
              fDtDateRevision := fDtLastDateRevCalc;

            fDtInitiale         := vQr.FindField('AFC_DATEINITIALE').AsDateTime;
            fDtPremiereRev      := vQr.FindField('AFC_PREMIEREDATE').AsDateTime;
            fBoAppliqCoefAuto   := vQr.FindField('AFC_APPLIQUERCOEF').AsString = 'X';
            fStPeriodeRev       := vQr.FindField('AFC_PERIODREV').AsString;
            fStModeLecture      := vQr.FindField('AFC_MODELECTURE').AsString;
            fInMoisLecture      := vQr.FindField('AFC_NUMMOISREV').AsInteger;
            fBoFinMois          := vQr.FindField('AFC_REVFINMOIS').AsString = 'X';
            fInJourLecture      := vQr.FindField('AFC_NUMJOURREV').AsInteger;
            fRdSeuilMini        := vQr.FindField('AFC_SEUILMIN').AsFloat;
            fRdSeuilMaxi        := vQr.FindField('AFC_SEUILMAX').AsFloat;
            fBoParamFormuleOK   := vQr.FindField('AFC_PARAMFORMULEOK').AsString = 'X';
          end;
        end
        else
          result := false;
      end
      else
        result := false;
    Finally
      if vQR <> nil then ferme(vQR);
    End;
  except
    result := false;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 02/04/2003
Modifié le ... :
Description .. : Retourne la valeur d'un indice
                 On se limite a 10 coefficients de passage
                 Il n'y a pas de limite de raccordement

                 la valeur des indices depend des parametres suivants :
                 pBoLastIndice : indique si on veut l'indice precedent ou l'indice suivant
                 pBoDatePublication : indique si on utilise la date de publication ou la date d'application (depend du type de l'indice)
Mots clefs ... :
*****************************************************************}
Function TCALCULCOEF.ValeurIndice(pDate : TDateTime; pStIndice, pStPubCode : String;
                                  pBoLastIndice, pBoDeuxDates, pBoInitial : Boolean; var pDtIndice,
                                  pDtIndDateVal : TDateTime; var pStIndPubCode : String;
                                  var pStIndiceErreur : String) : Double;
var
  i               : Integer;
  j               : Integer;
  vSt             : String;
  vQr             : TQuery;
  vBoTrouve       : Boolean;
  vStIndice       : String;
  vStPubCode      : String;
  vStCommun       : String;
  vRdRaccord      : Double;
  vQrRaccord      : TQuery;
  vTobRaccord     : Tob;

begin

  i       := 0;
  result  := -1;

  vBoTrouve   := False;
  vStIndice   := pStIndice;
  vStPubCode  := pStPubCode;

  // on se limite a 10 pour l'instant
  while not vBoTrouve and (i < 10) do
  begin
    // selon la méthode de calcul utilisée
    // on utilise la date de publication en respectant la date d'application
    // ou la date d'application
    if pBoDeuxDates then
      vSt := 'SELECT AFV_INDDATEPUB AS LADATE , '
    else
      vSt := 'SELECT AFV_INDDATEVAL AS LADATE , ';

    vSt := vSt + ' AFV_PUBCODE, AFV_INDDATEVAL, AFV_INDVALEUR, AFV_INDDATEFIN, ';
    vSt := vSt + ' AFV_INDCODESUIV, AFV_PUBCODESUIV ';
    vSt := vSt + ' FROM AFVALINDICE ';

    if pBoLastIndice then
    begin
      if pBoDeuxDates then
      begin
        vStCommun := ' WHERE AFV_INDDATEPUB <= "' + UsDateTime(pDate) + '"';
        vStCommun := vStCommun + ' AND AFV_INDDATEVAL <= "' + UsDateTime(pDate) + '"';
      end
      else
        vStCommun := ' WHERE AFV_INDDATEVAL <= "' + UsDateTime(pDate) + '"';
    end
    else
    begin
      if pBoDeuxDates then
      begin
        vStCommun := ' WHERE AFV_INDDATEPUB > "' + UsDateTime(pDate) + '"';
        vStCommun := vStCommun + ' AND AFV_INDDATEVAL > "' + UsDateTime(pDate) + '"';
      end
      else
        vStCommun := vStCommun + ' WHERE AFV_INDDATEVAL > "' + UsDateTime(pDate) + '"';
    end;

    vStCommun := vStCommun + ' AND AFV_INDCODE = "' + vStIndice + '"';
    if vStPubCode <> '' then
      vStCommun := vStCommun + ' AND AFV_PUBCODE = "' + vStPubCode + '"';

    // si on calcul un coef de regularisation
    // on ne prend que les indices definitifs
    if fBoCoefRegul then
      vStCommun := vStCommun + ' AND AFV_DEFINITIF = "X"';

    if pBoLastIndice then
      vSt := vSt + vStCommun + ' ORDER BY LADATE DESC'
    else
      vSt := vST + vStCommun + ' ORDER BY LADATE';

    vQr := nil;
    Try
      vQR := OpenSql(vSt, True,-1,'',true);
      if Not vQR.Eof then
      begin

        // pas d'indice suivant -> ok
        if (vQR.FindField('AFV_INDCODESUIV').AsString = '') then
        begin
          vBotrouve := True;
          result := vQR.FindField('AFV_INDVALEUR').AsFloat;
        end
        // date de fin d'indice est dépassée
        // indice plus utilisé -> coefficient de passage
        else if (vQR.FindField('AFV_INDDATEFIN').AsDateTime < pDate) then
        begin
          vStIndice   := vQR.FindField('AFV_INDCODESUIV').AsString;
          vStPubCode  := vQR.FindField('AFV_PUBCODESUIV').AsString;
        end
        else
        // indice sera remplacé, mais plus tard
        begin
          vBotrouve := True;
          result := vQR.FindField('AFV_INDVALEUR').AsFloat;
        end;

        // Date de l'indice
        pDtIndice := vQR.FindField('LADATE').AsDateTime;
        pDtIndDateVal := vQR.FindField('AFV_INDDATEVAL').AsDateTime;
        pStIndPubCode := vQR.FindField('AFV_PUBCODE').AsString;
      end
      else
      begin
        result := -1;
        pStIndiceErreur := vStIndice;
        break;
      end;

      // 03/10/03 pas de coefficient de raccordement pour des indices initiaux lus
      // dans la base (pas saisi dans les paramètrage)
      if not pBoInitial then
      begin
        // coefficient de raccordement
        // recherche de tous les coefficients de raccordement pour une date <=
        // a la date de l'indice actuelle
        vSt := ' SELECT AFV_COEFRACCORD ';
        vSt := vSt + ' FROM AFVALINDICE ';
        vSt := vSt + vStCommun;
        vSt := vSt + ' AND AFV_COEFRACCORD <> 0';
                          
        // 03/10/03
        // on ne recherche les raccordements que jusqu'à la date initiale des indices
        if (fDtInitiale <> iDate2099) then
        begin
          if pBoDeuxDates then
          begin
            vSt := vSt + ' AND AFV_INDDATEPUB >= "' + UsDateTime(fDtInitiale) + '"';
            vSt := vSt + ' AND AFV_INDDATEVAL >= "' + UsDateTime(fDtInitiale) + '"';
          end
          else
            vSt := vSt + ' AND AFV_INDDATEVAL >= "' + UsDateTime(fDtInitiale) + '"';
        end;

        vQrRaccord := nil;
        vRdRaccord := 1;
        vTobRaccord := Tob.create('AFVALINDICE', nil, -1);
        Try
          vQrRaccord := OpenSql(vSt, True,-1,'',true);
          if Not vQrRaccord.Eof then
            vTobRaccord.LoadDetailDB('AFVALINDICE','','',vQrRaccord,False,True);
          for j := 0 to vTobRaccord.detail.count -1 do
            vRdRaccord := vRdRaccord * vTobRaccord.detail[j].getvalue('AFV_COEFRACCORD');

          // C.B pour le raccordement, on ne fait plus une division,
          // mais une multiplication demande du 01/10/03
          if vRdRaccord <> 0 then
            result := vQR.FindField('AFV_INDVALEUR').AsFloat * vRdRaccord;

        finally
          if vQrRaccord <> nil then ferme(vQrRaccord);
          vTobRaccord.Free;
        end;
      end; // end if not pBoInitial

    Finally
      if vQR <> nil then ferme(vQR);
    end;
    i := i + 1;
  end;
End;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 03/04/2003
Modifié le ... :
Description .. : retourne la derniere valeur de l'indice utilisé pour cette formule
                 dans cette affaire
                 exclu les indices utilisés lors des regularisations
                 on prend exclusivement ceux de la facturation
Mots clefs ... :
*****************************************************************}
Function TCALCULCOEF.ValeursIndiceUtilise(pInNbValeurs : Integer; pStIndice : String; var pArValeurs : Array of double) : Integer;
var
  vSt   : String;
  vQr   : TQuery;
  i     : Integer;
  vTob  : Tob;

begin

  result := 0;

  vSt := 'SELECT AFR_DATECALCCOEF, ';
  vSt := vSt + ' AFR_VALINDICE' + intToStr(RangIndice(pStIndice)) + ' AS INDICE ';
  vSt := vSt + ' FROM AFREVISION ';
  vSt := vSt + ' WHERE AFR_AFFAIRE = "' + fStAffaire + '"';
  vSt := vSt + ' AND AFR_FORCODE = "' + fStCodeFormule + '"';
  vSt := vSt + ' AND AFR_COEFREGUL = "-"';
  vSt := vSt + ' GROUP BY AFR_DATECALCCOEF, AFR_VALINDICE' + intToStr(RangIndice(pStIndice));
  vSt := vSt + ' ORDER BY AFR_DATECALCCOEF DESC';

  vTob := TOB.Create('AFPARAMFORMULE', nil, -1);
  vQr := nil;
  Try
    vQR := OpenSql(vSt, True,-1,'',true);
    if Not vQR.Eof then
    begin
      vTob.LoadDetailDB('AFREVISION','','',vQR,False,True);
      for i := 0 to pInNbValeurs - 1 do
      begin
        if assigned(vTob.detail[i]) then
        begin
          pArValeurs[i] := vTob.detail[i].getValue('INDICE');
          result := i + 1;
        end
        else
        begin
          result := i;
          break;
        end;
      end;
    end;
  Finally
    if vQR <> nil then ferme(vQR);
    vTob.free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 05/05/2003
Modifié le ... :
Description .. : retourne les dernieres valeurs de l'indice saisies en base
Mots clefs ... :
*****************************************************************}
Function TCALCULCOEF.ValeursIndiceSaisis(pInNbValeurs : Integer; pStIndice,  pStPubCode: String; var pArValeurs : Array of double) : Integer;
var
  vSt         : String;
  vQr         : TQuery;
  i           : Integer;
  vTob        : Tob;
  vBoTrouve   : Boolean;
  vRdVal      : Double;
  vStIndice   : String;
  vStPubCode  : String;
  vInDecalage : Integer;

  // on prend les x derniers mois
  // si plusieurs saisies dans le meme mois, on ne prend pas la
  // seconde valeur
  function GetValeurIndice(pTob : Tob; i : Integer) : Double;
  var
    MoisPrec, Year, Month, Day : Word;
  begin

    result := -1;
    if assigned(pTob.detail[i]) then
    begin

      MoisPrec := 0;
      if i > 0 then
      begin
        DecodeDate(pTob.detail[i-1].getvalue('AFV_INDDATEVAL'), Year, Month, Day);
        MoisPrec := Month;
        DecodeDate(pTob.detail[i].getvalue('AFV_INDDATEVAL'), Year, Month, Day);
      end;

      // cas 1 : premier passage
      if (i = 0) or
      // cas 2 : mois du précédent est différent du mois en cours
      (month <> MoisPrec) then
      begin
        // coefficient de raccordement
        if (pTob.detail[i].getValue('AFV_COEFRACCORD') <> '0') and
           (pTob.detail[i].getValue('AFV_COEFRACCORD') <> '') then
          result := pTob.detail[i].getValue('AFV_INDVALEUR') * pTob.detail[i].getValue('AFV_COEFRACCORD')
        else
          result := pTob.detail[i].getValue('AFV_INDVALEUR');
      end;
    end;
  end;

begin

  result      := 0;
  i           := 0;
  vStIndice   := pStIndice;
  vStPubCode  := pStPubCode;
  vBotrouve   := False;

  // on se limite a 10 pour l'instant
  while not vBoTrouve and (i < 10) do
  begin

    vSt := 'SELECT AFV_INDDATEVAL, ';
    vSt := vSt + ' AFV_INDVALEUR, AFV_INDDATEFIN, AFV_INDCODESUIV, AFV_PUBCODESUIV, AFV_COEFRACCORD ';
    vSt := vSt + ' FROM AFVALINDICE ';
    vSt := vSt + ' WHERE AFV_INDCODE = "' + vStIndice + '"';
    if vStPubCode <> '' then
      vSt := vSt + ' AND AFV_PUBCODE = "' + vStPubCode + '"';

    vSt := vSt + ' AND AFV_INDDATEVAL <= "' + usDateTime(fDtDateRevision) + '"';

    vSt := vSt + ' GROUP BY AFV_INDDATEVAL, AFV_INDVALEUR, AFV_INDDATEFIN, AFV_INDCODESUIV, AFV_PUBCODESUIV, AFV_COEFRACCORD';
    vSt := vSt + ' ORDER BY AFV_INDDATEVAL DESC';

    vTob := TOB.Create('AFVALINDICE', nil, -1);
    vQr := nil;
    Try
      vQR := OpenSql(vSt, True,-1,'',true);
      if Not vQR.Eof then
      begin

        // pas d'indice suivant -> ok
        if (vQR.FindField('AFV_INDCODESUIV').AsString = '') then
        begin
          vBotrouve := True;

          vTob.LoadDetailDB('AFVALINDICE','','',vQR,False,True);
          i := 0;
          vInDecalage := 0;
          while i < pInNbValeurs + vInDecalage do
          begin
            if vTob.detail.count > i then
            begin
              vRdVal := GetValeurIndice(vTob, i);
              // si 2 valeurs pour le meme mois
              // on passe a la donnée suivante
              if vRdVal <> -1 then
              begin
                pArValeurs[i - vInDecalage] := vRdVal;
                result := i + 1;
              end
              else
                vInDecalage := vInDecalage + 1;
              i := i + 1;
            end
            else
            begin
              result := i;
              break;
            end;
          end;
        end

        // date de fin d'indice est dépassée
        // indice plus utilisé -> coefficient de passage
        else if (vTob.detail[0].GetValue('AFV_INDDATEFIN') < vTob.detail[0].GetValue('AFV_INDDATEVAL')) then
        begin
          vStIndice   := vTob.detail[0].GetValue('AFV_INDCODESUIV');
          vStPubCode  := vTob.detail[0].GetValue('AFV_PUBCODESUIV');
        end
        else
        // indice sera remplacé, mais plus tard
        begin
          vBotrouve := True;
          vTob.LoadDetailDB('AFVALINDICE','','',vQR,False,True);
          for i := 0 to pInNbValeurs - 1 do
          begin
            vRdVal := GetValeurIndice(vTob, i);
            if vRdVal = -1 then
            begin
              result := i;
              break;
            end
            else
            begin
              pArValeurs[i] := vRdVal;
              result := i + 1;
            end;
          end;
        end;
      end
      else                 
        result := 0;

    Finally
      if vQR <> nil then ferme(vQR);
      vTob.Free;
    end;
    i := i + 1;
  end;
End;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. : Recherche du rang dans la formule de l'indice recherché
Mots clefs ... :
*****************************************************************}
function TCALCULCOEF.RangIndice(pStIndice : String) : Integer;
var
  i : Integer;
begin
  result := 1;
  for i := 1 to fInNbindices do
    if fArCodeIndice[i] = pStIndice then
    begin
      result := i;
      break;
    end;
end;           

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 09/04/2003
Modifié le ... :
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TCALCULCOEF.CalculNbIndices;
begin
  fInNbindices := 0;
  while fArCodeIndice[fInNbindices + 1] <> '' do
    fInNbindices := fInNbindices + 1;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. :Si la date initiale est renseignée, le tableau n'est pas
                forcement remplit mais si il est remplit,
                il est prioritaire sur la date de debut
                Si la date initiale n'est pas renseignée,
                le tableau est obligatoirement renseigné

                indices chaines
                on prend la derniere valeur d'indice utilisé lors de la derniere revision
                comme indice 0. si on ne trouve rien, on prend les indices initiaux

Mots clefs ... :
*****************************************************************}
Function TCALCULCOEF.LoadValeurInitIndices(var pStIndiceErreur : String) : Boolean;
var               
  i             : Integer;
  vDate         : TDateTime;
  vInNbRetour   : Integer;
  vArValeurs    : Array[0 .. 9] of double;
  vDtIndDateVal : TDateTime;
  vStIndPubCode : String;

begin

  result := true;

  // recherche des valeurs des indices pour les valeurs non renseignées
  if fDtInitiale <> iDate1900 then
  begin
    for i := 1 to fInNbindices do
    begin

      // chainés
      // on va le chercher dans le dernier appliqué dans les lignes d'affaires
      // pas dans la facture
      if (fStTypeIndice = 'CHA') Then
      begin
        // dans ce cas, on ne se sert que de la valeur 0 du tableau
        vArValeurs[0] := 0;

        vInNbRetour := ValeursIndiceUtilise(1, fArCodeIndice[i], vArValeurs);
        Case vInNbRetour of
        // si pas de valeur précédente
        // on va chercher les indices initiaux
          0 : begin
                // si pas dans paramformule
                if fArValInitIndice[i] = 0 then
                begin
                  // lecture en base
                  fArValInitIndice[i] := ValeurIndice(fDtInitiale, fArCodeIndice[i], fArPubCode[i], cBoDernierIndice, cBoDtApplication, True {idice initial}, vDate, vDtIndDateVal, vStIndPubCode,pStIndiceErreur);
                  if fArValInitIndice[i] = -1 then
                  begin
                    result := false;
                    break;
                  end;
                end;
              end;
          1 : begin
                fArValInitIndice[i] := vArValeurs[0];
              end;
          else
            begin
              result := false;
              break;
            end;
        end;
      end

      // autres cas
      else
        if fArValInitIndice[i] = 0 then
        begin
          fArValInitIndice[i] := ValeurIndice(fDtInitiale, fArCodeIndice[i], fArPubCode[i], cBoDernierIndice, cBoDtApplication, True{indice initial}, vDate, vDtIndDateVal, vStIndPubCode, pStIndiceErreur);
          if fArValInitIndice[i] = -1 then
          begin
            result := false;
            break;
          end;
        end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 01/04/2003
Modifié le ... :
Description .. : En fonction du mode de lecture des indices,
                 on determine la date de lecture des indices pour
                 calculer le coefficient
                 On peut calculer des indices par rapport a la date
                 de facturation ou par rapport a la date de revision
                 on peut recalculer la derniere revision ou calculer
                 la prochaine revision
Mots clefs ... :
*****************************************************************}
function TCALCULCOEF.DateLectureIndices : TDateTime;
var
  vInMoisLecture : Integer;

  procedure IndiceDateRevision(pInMoisLecture : Integer);
  begin
    if fBoFinMois then
    begin
      if fBoProchaine then
        result := FinDeMois(PlusMois(fDtNextDateApp, vInMoisLecture))
      else
        result := FinDeMois(PlusMois(fDtLastDateRevCalc, vInMoisLecture))
    end
    else
    begin
      if fBoProchaine or (fDtLastDateRevApp = iDate2099) then
        result := PlusDate(PlusMois(DebutDeMois(fDtNextDateApp), vInMoisLecture), fInJourLecture -1, 'J')
      else
        result := PlusDate(PlusMois(DebutDeMois(fDtLastDateRevCalc), vInMoisLecture), fInJourLecture -1,  'J')
    end;
  end;

begin
  result := iDate2099;

  if fInMoisLecture > 0 then
    vInMoisLecture := fInMoisLecture -1
  else
    vInMoisLecture := fInMoisLecture;

  // on calcul un coef de regularisation
  if fBoCoefRegul then
  begin
    // lecture des indices par rapport a la facturation
    if fStModeLecture = 'MFA' then
    begin
      if fBoFinMois then
        result := FinDeMois(PlusMois(fDtEcheance, vInMoisLecture))
      else
        result := PlusDate(PlusMois(DebutDeMois(fDtEcheance), vInMoisLecture), fInJourLecture -1, 'J');
    end
    // lecture des indices par rapport a la revision
    else if fStModeLecture = 'MRE' then
      IndiceDateRevision(vInMoisLecture);
  end
  // on calcule un coef de revision
  else                                 
    IndiceDateRevision(vInMoisLecture);

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 02/04/2003
Modifié le ... :
Description .. : Calcule les valeurs des indices en fonction des types
                 de lecture des indices
Mots clefs ... :
*****************************************************************}
function TCALCULCOEF.LoadValeursIndices(var pStIndiceErreur : String) : Boolean;
var
  i, j, k           : Integer;
  vArValeurs        : Array[0 .. 9] of double;
  vRdValeur         : Double;
  vRdLastIndice     : Double;
  vRdNextIndice     : Double;
  vDateLast         : TDateTime;
  vDateNext         : TDateTime;
  vInNbJoursAvant   : Integer;
  vInNbJoursApres   : Integer;
  vInNbRetour       : Integer;
  vDtIndDateVal     : TDateTime;
  vStIndPubCode     : String;

Begin
  result := true;
  fDtLectureIndice := DateLectureIndices;
  // on prend la derniere valeur connue à la date de lecture d'indice / date d'application
  if (fStTypeIndice = 'REE') then
  begin
    for i := 1 to fInNbindices do
    begin
      fArValIndice[i]  := ValeurIndice(fDtLectureIndice, fArCodeIndice[i], fArPubCode[i], cBoDernierIndice, cBoDtApplication, False {pas indice initial}, vDateLast, vDtIndDateVal, vStIndPubCode, pStIndiceErreur);
      fArDatIndice[i]  := vDateLast;
      fArIndPubCode[i] := vStIndPubCode;
      fArIndDateVal[i] := vDtIndDateVal;

      if fArValIndice[i] = -1 then
      begin
        result := false;
        break;
      end;
    end;
  end

  // on prend la derniere valeur connue à la date de lecture d'indice / date de publication
  // cb 09/09/03 test sur la date de publication et la date d'application
  else if (fStTypeIndice = 'COS') Then
  begin
    for i := 1 to fInNbindices do
    begin
      fArValIndice[i] := ValeurIndice(fDtLectureIndice, fArCodeIndice[i], fArPubCode[i], cBoDernierIndice, cBoDeuxDates, False {pas indice initial}, vDateLast, vDtIndDateVal, vStIndPubCode, pStIndiceErreur);
      fArDatIndice[i] := vDateLast;
      fArIndPubCode[i] := vStIndPubCode;
      fArIndDateVal[i] := vDtIndDateVal;

      if fArValIndice[i] = -1 then
      begin
        result := false;
        break;
      end;
    end;
  end

  // chainés
  else if (fStTypeIndice = 'CHA') Then
  begin
    for i := 1 to fInNbindices do
    begin
      // valeurs indices
      fArValIndice[i] := ValeurIndice(fDtLectureIndice, fArCodeIndice[i], fArPubCode[i], cBoDernierIndice, cBoDtApplication, False {pas indice initial}, vDateLast,  vDtIndDateVal, vStIndPubCode, pStIndiceErreur);
      fArDatIndice[i] := vDateLast;
      fArIndPubCode[i] := vStIndPubCode;
      fArIndDateVal[i] := vDtIndDateVal;

      if fArValIndice[i] = -1 then
      begin
        result := false;
        break;
      end;
    end;
  end

  // calcul de la moyenne des x dernieres revisions
  // si l'historique est trop court, on fait la moyenne
  // sur les mois connus
  else if (fStTypeIndice = 'MOY') Then
  begin
    for i := 1 to fInNbindices do
    begin

      for j := 0 to fInNbindices -1 do
        vArValeurs[j] := 0;

      vInNbRetour := ValeursIndiceSaisis(fInNbMois, fArCodeIndice[i], fArPubCode[i], vArValeurs);

      if vInNbRetour = 0 then
      begin
        result := false;
        break;
      end
      else
      begin
        vRdValeur := 0;
        for k := 0 to vInNbRetour -1 do
          vRdValeur := vRdValeur + vArValeurs[k];
        fArValIndice[i] := vRdValeur / vInNbRetour;
      end;
    end;
  end
 
  // prorata temporis
  // on recherche l'indice précédent et l'indice suivant
  // et on proratise par rapport au nombre de jours
  else if (fStTypeIndice = 'PRO') Then
  begin
    for i := 1 to fInNbindices do
    begin
      vRdLastIndice := ValeurIndice(fDtLectureIndice, fArCodeIndice[i], fArPubCode[i], cBoDernierIndice, cBoDtApplication, False {pas indice initial}, vDateLast, vDtIndDateVal, vStIndPubCode, pStIndiceErreur);
      vRdNextIndice := ValeurIndice(fDtLectureIndice, fArCodeIndice[i], fArPubCode[i], cBoProchainIndice, cBoDtApplication, False {pas indice initial}, vDateNext, vDtIndDateVal, vStIndPubCode, pStIndiceErreur);
 
      if (vRdLastIndice = -1) or (vRdNextIndice = -1) then
      begin
        fArValIndice[i] := -1;                                                     
        result := false;
      end
      else
      begin
        vInNbJoursAvant := trunc(fDtLectureIndice - vDateLast);
        vInNbJoursApres := trunc(vDateNext - fDtLectureIndice);

        if (vInNbJoursAvant + vInNbJoursApres) <> 0 then
          fArValIndice[i] := ((vInNbJoursAvant * vRdLastIndice) + (vInNbJoursApres * vRdNextIndice)) / (vInNbJoursAvant + vInNbJoursApres)
        else
        begin
          fArValIndice[i] := -1;
          result := false;
        end;
      end;
    end;
  end
  else
    result := false;

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/04/2003
Modifié le ... :
Description .. : calcul du coefficient
Mots clefs ... :
*****************************************************************}
{function TCALCULCOEF.CalculCoef : boolean;
var
  i   : Integer;
  RF  : R_FonctCal;

begin

  RF := initRecord;
  RF.Formule := GFormuleVersFormatPolonais(fStFormule);
  RF.Affichage := fStFormule;
  for i := 1 to fInNbindices do
  begin
    RF.VarLibelle[i-1] := fArCodeIndice[i];
    RF.VarValeur[i-1] := fArValIndice[i];
    RF.VarLibelle[fInNbindices + i - 1] := fArCodeIndice[i] + '°';
    RF.VarValeur[fInNbindices + i - 1] := fArValInitIndice[i];
  end;

  RF := EvaluationFormule('', RF, True);
  fRdCoef := RF.Resultat;

end;
}

function TCALCULCOEF.CalculCoef : boolean;
var
  expr  : string;
  i     : Integer;

begin

  expr := FindEtReplace(fStFormule,'°','.0',true);
  for i := 1 to fInNbindices do
  begin
    expr := FindEtReplace(expr,'[' + fArCodeIndice[i] + '.0' + ']','[IND' + intToStr(i) + '.0]' , true);
    expr := FindEtReplace(expr,'[' + fArCodeIndice[i] + ']','[IND' + intToStr(i) + ']', true);
  end;
                             
  expr := FindEtReplace(expr,'[','',true);
  expr := FindEtReplace(expr,']','',true);
  expr := FindEtReplace(expr,'ARR{','@ARR(',true);
  expr := FindEtReplace(expr,'{','(',true);
  expr := FindEtReplace(expr,'}',')',true);
  expr := FindEtReplace(expr,';',',',true);
 
  result :=  ComputeRTFormule( expr, OnFunction, OnGetVar, fRdCoef) = 0 {COMPUTE_OK};

  // arrondi du coefficient
  // si la valeur est <> de 0
 if fInArrondiCoeff <> 0 then
    fRdCoef := strToFloat(format('%.' + intToStr(fInArrondiCoeff) + 'f', [fRdCoef]));

end;

//formule d'actualisation
function TCALCULCOEF.TestFormuleActualisation : boolean;
begin
  result := fStTypeFormule = 'ACT';
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 20/05/2003
Modifié le ... :   /  /
Description .. : Déclencher une actualisation
Mots clefs ... :
*****************************************************************}
function TCALCULCOEF.TestActualisation : boolean;
var
  vSt : String;
  vQr : TQuery;
  vDt : TDateTime;

begin
  // recherche si premiere facture
  vSt := 'SELECT GP_AFFAIRE FROM PIECE ';
  vSt := vSt + ' WHERE GP_AFFAIRE = "' + fStAffaire + '"';
  vSt := vSt + ' AND (GP_NATUREPIECEG = "FAC" OR GP_NATUREPIECEG = "FPR")';
  vQr := nil;
  Try
    vQR := OpenSql(vSt, True,-1,'',true);
    if vQR.Eof then
    begin
      // pas de facture
      // recherche si date -> actualisation
      ferme(vQR);
      vSt := 'SELECT AFF_DATEDEBUT, AFF_DATESIGNE FROM AFFAIRE';
      vSt := vSt + ' WHERE AFF_AFFAIRE = "' + fStAffaire + '"';
      vQr := nil;
      Try
        vQR := OpenSql(vSt, True,-1,'',true);
        if not vQR.Eof then
        begin
          // règle : date de début + 3 mois < date d'acceptation
          vDt := plusMois(vQr.FindField('AFF_DATEDEBUT').AsDateTime, 3);
          if vDt < vQr.FindField('AFF_DATESIGNE').AsDateTime then
            result := True
          else
            result := False;
        end 
        else
          result := false;
      Finally
        if vQR <> nil then ferme(vQR);
      End;
    end
    // l'actualisation a deja eu lieu
    else
      result := false;
  Finally
    if vQR <> nil then ferme(vQR);
  End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/04/2003
Modifié le ... :   /  /
Description .. : Sépare la formule de départ en autant de termes
                 qu'il y a d'arrondis
Mots clefs ... :
*****************************************************************}
{procedure TCALCULCOEF.DecodeTermes(pStFormule : String; var vArTermes, vArArrondis : Array of String);
var
  i               : Integer;
  vStFormule      : String;
  vInNbCaract     : Integer;
  vInNbCaractArr  : Integer;

begin
  i := 0;
  vStFormule := pStFormule;
  while pos(cStOuvreArrondi, vStFormule) > 0 do
  begin              
    // récupération données
    vInNbCaract := pos(cStSepArrondi, vStFormule) - 4 - pos(cStOuvreArrondi, vStFormule);
    vInNbCaractArr := pos(cStFermeArrondi, vStFormule) - pos(cStSepArrondi, vStFormule) -1;
    vArTermes[i] := copy(vStFormule, pos(cStOuvreArrondi, vStFormule) + 4, vInNbCaract);
    vArArrondis[i] := trim(copy(vStFormule, pos(cStSepArrondi, vStFormule) + 1, vInNbCaractArr));

    // supression du terme
    vStFormule := copy(vStFormule, pos(cStFermeArrondi, vStFormule) + 1, length(vStFormule));

    i := i + 1;
  end;

  // recuperation de la fin de la formule
  if length(vStFormule) > 0 then
  begin
    vArTermes[i] := vStFormule;
    vArArrondis[i] := '0';
  end;
end;
}

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/04/2003
Modifié le ... :   /  /
Description .. : Remplace les indices par leurs valeurs
Mots clefs ... :
*****************************************************************}
{function TCALCULCOEF.SetValIndiceFormule(pStIndice : String) : variant;
var
  i : Integer;
  vStIndice : String;

begin
  for i := 1 to fInNbindices do
  begin
    if IndiceZero(pStIndice) then
    begin
      vStIndice := pStIndice;
      ReplaceSubStr(vStIndice, cStIndiceZero, '');
      if vStIndice = fArCodeIndice[i] then
      begin
        result := fArValInitIndice[i];
        break;
      end
    end
    else
      if pStIndice = fArCodeIndice[i] then
      begin
        result := fArValIndice[i];
        break;
      end;
  end;
end;
}

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/04/2003
Modifié le ... :
Description .. : indique si c'est la valeur initiale de l'indice qui
                 doit etre utilisée
Mots clefs ... :
*****************************************************************}
{function TCALCULCOEF.IndiceZero(pStIndice : String) : Boolean;
begin
  result := pos(cStIndiceZero, pStIndice) > 0;
end;
}

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : Parcours d'une affaire et des lignes de l'affaire
                 pour créer toutes les lignes ayant la formule traitée
                 On supprime les enregistrements correspondants a cette
                 Régularisation
Mots clefs ... :
*****************************************************************}
Function TCALCULCOEF.InsertCoefRegul(pDateRegul : TDateTime) : Integer;
var
  i               : Integer;
  vQr             : TQuery;
  vSt             : String;
  vTob            : Tob;
  vTobInsert      : Tob;
  vTobFille       : Tob;
  vInCpt          : Integer;

begin

  result := 0;

  vSt := 'DELETE FROM AFREVISION WHERE AFR_AFFAIRE = "' + fStAffaire + '"';
  vSt := vSt + ' AND AFR_COEFREGUL = "X"';
  vSt := vSt + ' AND AFR_FORCODE = "' + fStCodeFormule + '"';
  vSt := vSt + ' AND AFR_DATECALCCOEF = "' + usDateTime(pDateRegul) + '"';
  vSt := vSt + ' AND AFR_OKCOEFAPPLIQUE = "X"';
  executeSql(vSt);

  vInCpt := MaxRevision;

  // Coef de l'affaire
  vSt := 'SELECT AFF_FORCODE1, AFF_FORCODE2 FROM AFFAIRE WHERE AFF_AFFAIRE = "' + fStAffaire + '"';
  vTobInsert := Tob.Create('Revision mere', nil, -1);
  vQr := nil;
  Try
    vQR := OpenSql(vSt, True,-1,'',true);
    if Not vQR.Eof then
    begin
      if (vQR.FindField('AFF_FORCODE1').AsString = fStCodeFormule) or
         (vQR.FindField('AFF_FORCODE2').AsString = fStCodeFormule) then
      begin
        vTobFille := Tob.Create('AFREVISION', vTobInsert, -1);

        // initialisation clé
        vTobFille.PutValue('AFR_AFFAIRE', fStAffaire);
        vTobFille.PutValue('AFR_NATUREPIECEG', cStVide);
        vTobFille.PutValue('AFR_SOUCHE', cStVide);
        vTobFille.PutValue('AFR_NUMERO', cInVide);
        vTobFille.PutValue('AFR_INDICEG', cInVide);
        vTobFille.PutValue('AFR_NUMORDRE', cInVide);
        // coef de regul applique toujours appliqués
        vTobFille.PutValue('AFR_OKCOEFAPPLIQUE', 'X');

        // champs de révision
        if (vQR.FindField('AFF_FORCODE1').AsString = fStCodeFormule) then
          AddLigneCoef(vTobFille, cInFormule1, vInCpt)
        else
          AddLigneCoef(vTobFille, cInFormule2, vInCpt);
      end;
    end;
  Finally
    if vQR <> nil then ferme(vQR);
  end;

  // on calcule egalement si la quantité facturé est 0, pour les lignes
  // qui sont transferées dans l'activité
  vSt := 'SELECT GL_NATUREPIECEG, GL_SOUCHE, GL_NUMERO, GL_INDICEG, ';
  vSt := vSt + ' GL_NUMORDRE, GL_FORCODE1, GL_FORCODE2 FROM LIGNE ';
  vSt := vSt + ' WHERE GL_AFFAIRE = "' + fStAffaire + '"';
  vSt := vSt + ' AND GL_NATUREPIECEG = "' + GetParamSoc('SO_AFNATAFFAIRE') + '"';

  vQr := nil;
  vTob := TOB.Create('LIGNE', nil, -1);
  Try
    vQR := OpenSql(vSt, True,-1,'',true);
    if Not vQR.Eof then
    begin
      vTob.LoadDetailDB('LIGNE','','',vQR,False,True);
      for i := 0 to vTob.detail.count - 1 do
      begin
        if (vTob.detail[i].GetValue('GL_FORCODE1') = fStCodeFormule) or
           (vTob.detail[i].GetValue('GL_FORCODE2') = fStCodeFormule) then
        begin
          vTobFille := Tob.Create('AFREVISION', vTobInsert, -1);

          // initialisation clé
          vTobFille.PutValue('AFR_AFFAIRE', fStAffaire);
          vTobFille.PutValue('AFR_NATUREPIECEG', vTob.Detail[i].GetValue('GL_NATUREPIECEG'));
          vTobFille.PutValue('AFR_SOUCHE', vTob.Detail[i].GetValue('GL_SOUCHE'));
          vTobFille.PutValue('AFR_NUMERO', vTob.Detail[i].GetValue('GL_NUMERO'));
          vTobFille.PutValue('AFR_INDICEG', vTob.Detail[i].GetValue('GL_INDICEG'));
          vTobFille.PutValue('AFR_NUMORDRE', vTob.Detail[i].GetValue('GL_NUMORDRE'));

          // table origine
          vTobFille.PutValue('AFR_ONATUREPIECEG', vTob.Detail[i].GetValue('GL_NATUREPIECEG'));
          vTobFille.PutValue('AFR_OSOUCHE', vTob.Detail[i].GetValue('GL_SOUCHE'));
          vTobFille.PutValue('AFR_ONUMERO', vTob.Detail[i].GetValue('GL_NUMERO'));
          vTobFille.PutValue('AFR_OINDICEG', vTob.Detail[i].GetValue('GL_INDICEG'));
          vTobFille.PutValue('AFR_ONUMORDRE', vTob.Detail[i].GetValue('GL_NUMORDRE'));

          // champs de révision
          if (vTob.detail[i].GetValue('GL_FORCODE1') = fStCodeFormule) then
            AddLigneCoef(vTobFille, cInFormule1, vinCpt)
          else
            AddLigneCoef(vTobFille, cInFormule2, vInCpt);
        end;
      end;
    end;

    // la révision est terminée
    if vTobInsert.InsertDB(nil, false) then
      fSLLog.Add(DateToStr(Now) + ' ' + format(TexteMsg[10], [fStAffaire, fStCodeFormule, fStLibFormule]))
    else
      result := -2;

  Finally
    if vQR <> nil then ferme(vQR);
    vTob.Free;
    vTobInsert.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : Parcours d'une affaire et des lignes de l'affaire
                 pour créer toutes les lignes ayant la formule traitée
                 Si on recalcule la derniere revision, on supprime
                 les enregistrements qui sont recalculés
Mots clefs ... :
*****************************************************************}
Function TCALCULCOEF.InsertCoefficients(pBoRecupTob : Boolean; pTobParamFormule, pTobRevision : Tob) : Integer;
var
  i               : Integer;
  vQr             : TQuery;
  vSt             : String;
  vTob            : Tob;
  vTobInsert      : Tob;
  vTobFille       : Tob;
  vTobLastCoef    : Tob;
  vTobTheLastCoef : Tob;
  vInCpt          : Integer;

begin

  result := 0;

  // suppression des éléments recalculés s'ils ne sont pas appliqués
  // si des coefficients sont deja appliqués, on empêche le recalcul

  if not fBoProchaine then
  begin
    vSt := 'SELECT AFR_AFFAIRE FROM AFREVISION ';
    vSt := vSt + ' WHERE AFR_AFFAIRE = "' + fStAffaire + '"';
    vSt := vSt + ' AND AFR_COEFREGUL = "-"';
    vSt := vSt + ' AND AFR_FORCODE = "' + fStCodeFormule + '"';
    vSt := vSt + ' AND AFR_DATECALCCOEF = "' + usDateTime(fDtLastDateRevCalc) + '"';
    vSt := vSt + ' AND AFR_OKCOEFAPPLIQUE = "X"';

    if ExisteSql(vSt) then
    begin
      fSLLog.add(DateToStr(now) + ' ' + format(TexteMsg[13], [fStAffaire, fStCodeFormule, fStLibFormule]));
      result := -1;
    end
    else
    begin
      vSt := 'DELETE FROM AFREVISION WHERE AFR_AFFAIRE = "' + fStAffaire + '"';
      vSt := vSt + ' AND AFR_COEFREGUL = "-"';
      vSt := vSt + ' AND AFR_FORCODE = "' + fStCodeFormule + '"';
      vSt := vSt + ' AND AFR_DATECALCCOEF = "' + usDateTime(fDtLastDateRevCalc) + '"';
      vSt := vSt + ' AND AFR_OKCOEFAPPLIQUE = "-"';
      executeSql(vSt);
    end;
  end;

  vInCpt := MaxRevision;

  // Chargement des anciens Coef de cette affaire
  vSt := 'SELECT AFR_AFFAIRE, AFR_NATUREPIECEG, AFR_SOUCHE, ';
  vSt := vSt + ' AFR_NUMERO, AFR_INDICEG, AFR_NUMORDRE, ';
  vSt := vSt + ' AFR_FORCODE, AFR_COEFCALC, AFR_COEFAPPLIQUE FROM AFREVISION ';
  vSt := vSt + ' WHERE AFR_AFFAIRE = "' + fStAffaire + '"';
  vSt := vSt + ' AND AFR_COEFREGUL = "-"' ;
  vSt := vSt + ' ORDER BY AFR_DATECALCCOEF DESC';

  vTobLastCoef := TOB.Create('AFREVISION', nil, -1);
  vQr := nil;
  Try
    vQR := OpenSql(vSt, True,-1,'',true);
    if Not vQR.Eof then
    begin
      vTobLastCoef.LoadDetailDB('AFREVISION','','',vQR,False,True);
    end;
    ferme(vQR);

    // Coef de l'affaire
    vSt := 'SELECT AFF_FORCODE1, AFF_FORCODE2 FROM AFFAIRE WHERE AFF_AFFAIRE = "' + fStAffaire + '"';
    vTobInsert := Tob.Create('Revision mere', nil, -1);
    vQr := nil;
    Try
      vQR := OpenSql(vSt, True,-1,'',true);
      if Not vQR.Eof then
      begin
        if (vQR.FindField('AFF_FORCODE1').AsString = fStCodeFormule) or
           (vQR.FindField('AFF_FORCODE2').AsString = fStCodeFormule) then
        begin
          vTobFille := Tob.Create('AFREVISION', vTobInsert, -1);

          // initialisation clé
          vTobFille.PutValue('AFR_AFFAIRE', fStAffaire);
          vTobFille.PutValue('AFR_NATUREPIECEG', cStVide);
          vTobFille.PutValue('AFR_SOUCHE', cStVide);
          vTobFille.PutValue('AFR_NUMERO', cInVide);
          vTobFille.PutValue('AFR_INDICEG', cInVide);
          vTobFille.PutValue('AFR_NUMORDRE', cInVide);

          // anciennes valeurs de coef
          vTobTheLastCoef := vTobLastCoef.FindFirst(['AFR_AFFAIRE',
                                                     'AFR_NATUREPIECEG',
                                                     'AFR_SOUCHE',
                                                     'AFR_NUMERO',
                                                     'AFR_INDICEG',
                                                     'AFR_NUMORDRE',
                                                     'AFR_FORCODE'],
                                                     [fStAffaire,
                                                     cStVide,
                                                     cStVide,
                                                     cInVide,
                                                     cInVide,
                                                     cInVide,
                                                     fStCodeFormule], True);

          if vTobTheLastCoef <> nil then
          begin
            vTobFille.PutValue('AFR_DERNIERCOEF', vTobTheLastCoef.GetValue('AFR_COEFCALC'));
            vTobFille.PutValue('AFR_DERNCOEFCALC', vTobTheLastCoef.GetValue('AFR_COEFAPPLIQUE'));
          end
          else
          begin
            vTobFille.PutValue('AFR_DERNIERCOEF', 1);
            vTobFille.PutValue('AFR_DERNCOEFCALC', 1);
          end;

          // champs de révision
          if (vQR.FindField('AFF_FORCODE1').AsString = fStCodeFormule) then
            AddLigneCoef(vTobFille, cInFormule1, vInCpt)
          else
            AddLigneCoef(vTobFille, cInFormule2, vInCpt);
        end;
      end;
    Finally
      if vQR <> nil then ferme(vQR);
    end;

    // on calcule egalement si la quantité facturé est 0, pour les lignes
    // qui sont transferées dans l'activité
    vSt := 'SELECT GL_NATUREPIECEG, GL_SOUCHE, GL_NUMERO, GL_INDICEG, ';
    vSt := vSt + ' GL_NUMORDRE, GL_FORCODE1, GL_FORCODE2 FROM LIGNE ';
    vSt := vSt + ' WHERE GL_AFFAIRE = "' + fStAffaire + '"';
    vSt := vSt + ' AND GL_NATUREPIECEG = "' + GetParamSoc('SO_AFNATAFFAIRE') + '"';
//    vSt := vSt + ' AND GL_QTEFACT <> 0';

    vQr := nil;
    vTob := TOB.Create('LIGNE', nil, -1);
    Try
      vQR := OpenSql(vSt, True,-1,'',true);
      if Not vQR.Eof then
      begin
        vTob.LoadDetailDB('LIGNE','','',vQR,False,True);
        for i := 0 to vTob.detail.count - 1 do
        begin
          if (vTob.detail[i].GetValue('GL_FORCODE1') = fStCodeFormule) or
             (vTob.detail[i].GetValue('GL_FORCODE2') = fStCodeFormule) then
          begin
            vTobFille := Tob.Create('AFREVISION', vTobInsert, -1);

            // initialisation clé
            vTobFille.PutValue('AFR_AFFAIRE', fStAffaire);
            vTobFille.PutValue('AFR_NATUREPIECEG', vTob.Detail[i].GetValue('GL_NATUREPIECEG'));
            vTobFille.PutValue('AFR_SOUCHE', vTob.Detail[i].GetValue('GL_SOUCHE'));
            vTobFille.PutValue('AFR_NUMERO', vTob.Detail[i].GetValue('GL_NUMERO'));
            vTobFille.PutValue('AFR_INDICEG', vTob.Detail[i].GetValue('GL_INDICEG'));
            vTobFille.PutValue('AFR_NUMORDRE', vTob.Detail[i].GetValue('GL_NUMORDRE'));

            // table origine
            vTobFille.PutValue('AFR_ONATUREPIECEG', vTob.Detail[i].GetValue('GL_NATUREPIECEG'));
            vTobFille.PutValue('AFR_OSOUCHE', vTob.Detail[i].GetValue('GL_SOUCHE'));
            vTobFille.PutValue('AFR_ONUMERO', vTob.Detail[i].GetValue('GL_NUMERO'));
            vTobFille.PutValue('AFR_OINDICEG', vTob.Detail[i].GetValue('GL_INDICEG'));
            vTobFille.PutValue('AFR_ONUMORDRE', vTob.Detail[i].GetValue('GL_NUMORDRE'));


            // anciennes valeurs de coef
            vTobTheLastCoef := vTobLastCoef.FindFirst(['AFR_AFFAIRE',
                                                       'AFR_NATUREPIECEG',
                                                       'AFR_SOUCHE',
                                                       'AFR_NUMERO',
                                                       'AFR_INDICEG',
                                                       'AFR_NUMORDRE'],
                                                       [fStAffaire,
                                                       vTob.Detail[i].GetValue('GL_NATUREPIECEG'),
                                                       vTob.Detail[i].GetValue('GL_SOUCHE'),
                                                       vTob.Detail[i].GetValue('GL_NUMERO'),
                                                       vTob.Detail[i].GetValue('GL_INDICEG'),
                                                       vTob.Detail[i].GetValue('GL_NUMORDRE')], True);

            if vTobTheLastCoef <> nil then
            begin
              vTobFille.PutValue('AFR_DERNIERCOEF', vTobTheLastCoef.GetValue('AFR_COEFCALC'));
              vTobFille.PutValue('AFR_DERNCOEFCALC', vTobTheLastCoef.GetValue('AFR_COEFAPPLIQUE'));
            end
            else
            begin
              vTobFille.PutValue('AFR_DERNIERCOEF', 1);
              vTobFille.PutValue('AFR_DERNCOEFCALC', 1);
            end;

            // champs de révision
            if (vTob.detail[i].GetValue('GL_FORCODE1') = fStCodeFormule) then
              AddLigneCoef(vTobFille, cInFormule1, vinCpt)
            else
              AddLigneCoef(vTobFille, cInFormule2, vInCpt);
          end;
        end;
      end;

      // la révision est terminée
      if vTobInsert.InsertDB(nil, false) then
      begin
        // pour la facture en transaction on retourne les données en tob
        if pBoRecupTob then
          pTobRevision.Dupliquer(vTobInsert, True, True);

        if (fBoProchaine or fBoAppliqCoefAuto) then
          MajDatesRevisionsApresApplication(false, pBoRecupTob, pTobParamFormule);

        if fBoProchaine then
          fSLLog.Add(DateToStr(Now) + ' ' + format(TexteMsg[9], [fStAffaire, fStCodeFormule, fStLibFormule]))
        else
          fSLLog.Add(DateToStr(Now) + ' ' + format(TexteMsg[10], [fStAffaire, fStCodeFormule, fStLibFormule]));
      end
      else
        result := -2;

    Finally
      if vQR <> nil then ferme(vQR);
      vTob.Free;
      vTobInsert.Free;
    end;

  Finally
    vTobLastCoef.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/04/2003
Modifié le ... :   /  /
Description .. : le max est toute formule confondue
Mots clefs ... :
*****************************************************************}
function TCALCULCOEF.MaxRevision : Integer;
var
  vSt : String;
  vQr : TQuery;

Begin
  vSt := 'SELECT MAX(AFR_NUMEROLIGNE) AS NUM FROM AFREVISION ';
  vSt := vSt + 'WHERE AFR_AFFAIRE = "' + fStAffaire + '"';

  vQr := nil;
  Try
    vQR := OpenSql(vSt, True,-1,'',true);
    if Not vQR.Eof then
      result := vQr.FindField('NUM').AsInteger + 1
    else
      result := 1;
  Finally
    if vQR <> nil then ferme(vQR);
  End;
End;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 07/04/2003
Modifié le ... :   /  /
Description .. : Ajout des champs de revision dans la ligne de revision
                 a inserer
Mots clefs ... :
*****************************************************************}
procedure TCALCULCOEF.AddLigneCoef(pTob : Tob; pInNumFormule : Integer; var pInCpt : Integer);
Begin

  if fBoCoefRegul then
    pTob.PutValue('AFR_DATECALCCOEF', fDtEcheance)
  else
    pTob.PutValue('AFR_DATECALCCOEF', fDtDateRevision);
  pTob.PutValue('AFR_FORCODE', fStCodeFormule);
  pTob.PutValue('AFR_NUMFORMULE', pInNumFormule);
  pTob.PutValue('AFR_COEFCALC', fRdCoef);
  pTob.PutValue('AFR_COEFAPPLIQUE', fRdCoef);
  pTob.PutValue('AFR_NUMEROLIGNE', pInCpt);

  if fBoAppliqCoefAuto then
  begin
    pTob.PutValue('AFR_APPLIQUECOEF', 'X');
    pTob.PutValue('AFR_OKCOEFAPPLIQUE', 'X');
  end
  else
  begin
    pTob.PutValue('AFR_APPLIQUECOEF', '-');
    pTob.PutValue('AFR_OKCOEFAPPLIQUE', '-');
  end;
                   
  // laisser apres if fBoAppliqCoefAuto
  // C.B 16/09/03
  if fBoCoefRegul then
  begin
    pTob.PutValue('AFR_COEFREGUL', 'X');
    pTob.PutValue('AFR_OKCOEFAPPLIQUE', 'X');
  end
  else
    pTob.PutValue('AFR_COEFREGUL', '-');


  pTob.PutValue('AFR_VALINDICE1', fArValIndice[1]);
  pTob.PutValue('AFR_VALINDICE2', fArValIndice[2]);
  pTob.PutValue('AFR_VALINDICE3', fArValIndice[3]);
  pTob.PutValue('AFR_VALINDICE4', fArValIndice[4]);
  pTob.PutValue('AFR_VALINDICE5', fArValIndice[5]);
  pTob.PutValue('AFR_VALINDICE6', fArValIndice[6]);
  pTob.PutValue('AFR_VALINDICE7', fArValIndice[7]);
  pTob.PutValue('AFR_VALINDICE8', fArValIndice[8]);
  pTob.PutValue('AFR_VALINDICE9', fArValIndice[9]);
  pTob.PutValue('AFR_VALINDICE10',fArValIndice[10]);

  pTob.PutValue('AFR_VALINIT1', fArValInitIndice[1]);
  pTob.PutValue('AFR_VALINIT2', fArValInitIndice[2]);
  pTob.PutValue('AFR_VALINIT3', fArValInitIndice[3]);
  pTob.PutValue('AFR_VALINIT4', fArValInitIndice[4]);
  pTob.PutValue('AFR_VALINIT5', fArValInitIndice[5]);
  pTob.PutValue('AFR_VALINIT6', fArValInitIndice[6]);
  pTob.PutValue('AFR_VALINIT7', fArValInitIndice[7]);
  pTob.PutValue('AFR_VALINIT8', fArValInitIndice[8]);
  pTob.PutValue('AFR_VALINIT9', fArValInitIndice[9]);
  pTob.PutValue('AFR_VALINIT10',fArValInitIndice[10]);

  pTob.PutValue('AFR_PUBCODE1', fArIndPubCode[1]);
  pTob.PutValue('AFR_PUBCODE2', fArIndPubCode[2]);
  pTob.PutValue('AFR_PUBCODE3', fArIndPubCode[3]);
  pTob.PutValue('AFR_PUBCODE4', fArIndPubCode[4]);
  pTob.PutValue('AFR_PUBCODE5', fArIndPubCode[5]);
  pTob.PutValue('AFR_PUBCODE6', fArIndPubCode[6]);
  pTob.PutValue('AFR_PUBCODE7', fArIndPubCode[7]);
  pTob.PutValue('AFR_PUBCODE8', fArIndPubCode[8]);
  pTob.PutValue('AFR_PUBCODE9', fArIndPubCode[9]);
  pTob.PutValue('AFR_PUBCODE10',fArIndPubCode[10]);


  if fArIndDateVal[1] <> 0 then pTob.PutValue('AFR_INDDATEVAL1', fArIndDateVal[1]) else pTob.PutValue('AFR_INDDATEVAL1', iDate1900);
  if fArIndDateVal[2] <> 0 then pTob.PutValue('AFR_INDDATEVAL2', fArIndDateVal[2]) else pTob.PutValue('AFR_INDDATEVAL2', iDate1900);
  if fArIndDateVal[3] <> 0 then pTob.PutValue('AFR_INDDATEVAL3', fArIndDateVal[3]) else pTob.PutValue('AFR_INDDATEVAL3', iDate1900);
  if fArIndDateVal[4] <> 0 then pTob.PutValue('AFR_INDDATEVAL4', fArIndDateVal[4]) else pTob.PutValue('AFR_INDDATEVAL4', iDate1900);
  if fArIndDateVal[5] <> 0 then pTob.PutValue('AFR_INDDATEVAL5', fArIndDateVal[5]) else pTob.PutValue('AFR_INDDATEVAL5', iDate1900);
  if fArIndDateVal[6] <> 0 then pTob.PutValue('AFR_INDDATEVAL6', fArIndDateVal[6]) else pTob.PutValue('AFR_INDDATEVAL6', iDate1900);
  if fArIndDateVal[7] <> 0 then pTob.PutValue('AFR_INDDATEVAL7', fArIndDateVal[7]) else pTob.PutValue('AFR_INDDATEVAL7', iDate1900);
  if fArIndDateVal[8] <> 0 then pTob.PutValue('AFR_INDDATEVAL8', fArIndDateVal[8]) else pTob.PutValue('AFR_INDDATEVAL8', iDate1900);
  if fArIndDateVal[9] <> 0 then pTob.PutValue('AFR_INDDATEVAL9', fArIndDateVal[9]) else pTob.PutValue('AFR_INDDATEVAL9', iDate1900);
  if fArIndDateVal[10] <> 0 then pTob.PutValue('AFR_INDDATEVAL10', fArIndDateVal[10]) else pTob.PutValue('AFR_INDDATEVAL10', iDate1900);

  if fArDatIndice[1] <> 0 then pTob.PutValue('AFR_DATEINDICE1', fArDatIndice[1]) else pTob.PutValue('AFR_DATEINDICE1', iDate1900);
  if fArDatIndice[2] <> 0 then pTob.PutValue('AFR_DATEINDICE2', fArDatIndice[2]) else pTob.PutValue('AFR_DATEINDICE2', iDate1900);
  if fArDatIndice[3] <> 0 then pTob.PutValue('AFR_DATEINDICE3', fArDatIndice[3]) else pTob.PutValue('AFR_DATEINDICE3', iDate1900);
  if fArDatIndice[4] <> 0 then pTob.PutValue('AFR_DATEINDICE4', fArDatIndice[4]) else pTob.PutValue('AFR_DATEINDICE4', iDate1900);
  if fArDatIndice[5] <> 0 then pTob.PutValue('AFR_DATEINDICE5', fArDatIndice[5]) else pTob.PutValue('AFR_DATEINDICE5', iDate1900);
  if fArDatIndice[6] <> 0 then pTob.PutValue('AFR_DATEINDICE6', fArDatIndice[6]) else pTob.PutValue('AFR_DATEINDICE6', iDate1900);
  if fArDatIndice[7] <> 0 then pTob.PutValue('AFR_DATEINDICE7', fArDatIndice[7]) else pTob.PutValue('AFR_DATEINDICE7', iDate1900);
  if fArDatIndice[8] <> 0 then pTob.PutValue('AFR_DATEINDICE8', fArDatIndice[8]) else pTob.PutValue('AFR_DATEINDICE8', iDate1900);
  if fArDatIndice[9] <> 0 then pTob.PutValue('AFR_DATEINDICE9', fArDatIndice[9]) else pTob.PutValue('AFR_DATEINDICE9', iDate1900);
  if fArDatIndice[10] <> 0 then pTob.PutValue('AFR_DATEINDICE10', fArDatIndice[10]) else pTob.PutValue('AFR_DATEINDICE10', iDate1900);

  pInCpt := pInCpt + 1;
End;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 01/04/2003
Modifié le ... :
Description .. : mise a jour des dates de revision
                 utilisée en fin de revision apres la creation des enregistrements
Mots clefs ... :
*****************************************************************}
function TCALCULCOEF.MajDatesRevisionsApresApplication(pBoForcerApplication, pBoRecupTob : boolean; pTobParamFormule : Tob) : boolean;
var
  vSt     : string;
  vDtRef  : TDateTime;

begin
  // Dates vides -> premiere revision
  if (fDtNextDateRevCalc = iDate2099) then
    vDtRef := fDtPremiereRev

  // on calcule les nouvelles dates de revision à partir de l'ancienne date de revision
  else
    vDtRef := fDtNextDateApp;

  if fBoAppliqCoefAuto or pBoForcerApplication then
    fDtLastDateRevApp   := vDtRef;

  fDtLastDateRevCalc  := vDtRef;

  if fStPeriodeRev = 'A' then
    fDtNextDateRevCalc := PlusMois(vDtRef, 12)

  else if fStPeriodeRev = 'H' then
    fDtNextDateRevCalc := PlusDate(vDtRef, 1, 'S')

  else if fStPeriodeRev = 'M' then
    fDtNextDateRevCalc := PlusMois(vDtRef, 1)

  else if fStPeriodeRev = 'S' then
    fDtNextDateRevCalc := PlusMois(vDtRef, 6)

  else if fStPeriodeRev = 'T' then
    fDtNextDateRevCalc := PlusMois(vDtRef, 3);

  if fBoAppliqCoefAuto or pBoForcerApplication then
    fDtNextDateApp := fDtNextDateRevCalc
  else
    fDtNextDateApp := fDtLastDateRevCalc;

  vSt := 'UPDATE AFPARAMFORMULE SET AFC_LASTDATEAPP = "' + usdatetime(fDtLastDateRevApp) + '",';
  vSt := vSt + ' AFC_LASTDATECALC = "' + usdatetime(fDtLastDateRevCalc) + '",';
  vSt := vSt + ' AFC_NEXTDATECALC = "' + usdatetime(fDtNextDateRevCalc) + '",';
  vSt := vSt + ' AFC_NEXTDATEAPP = "' + usdatetime(fDtNextDateApp) + '"';
  vSt := vSt + ' WHERE AFC_AFFAIRE = "' + fStAffaire + '"';
  vSt := vSt + ' AND AFC_FORCODE = "' + fStCodeFormule + '"';

  if executeSql(vSt) = 1 then
    result := true
  else
    result := false;

  // retour des données sous forme de tob
  if pBoRecupTob then
  begin
    PtobParamFormule.AddChampSupValeur('AFC_AFFAIRE', fStAffaire);
    PtobParamFormule.AddChampSupValeur('AFC_FORCODE', fStCodeFormule);
    PtobParamFormule.AddChampSupValeur('AFC_LASTDATEAPP', fDtLastDateRevApp);
    PtobParamFormule.AddChampSupValeur('AFC_LASTDATECALC', fDtLastDateRevCalc);
    PtobParamFormule.AddChampSupValeur('AFC_NEXTDATECALC', fDtNextDateRevCalc);
    PtobParamFormule.AddChampSupValeur('AFC_NEXTDATEAPP', fDtNextDateApp);
    PtobParamFormule.AddChampSupValeur('AFC_APPLIQUERCOEF', fDtNextDateRevCalc);
    PtobParamFormule.AddChampSupValeur('AFC_PERIODREV', fDtNextDateApp);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 01/04/2003
Modifié le ... :
Description .. : Calcul du coefficient pour une formule d'une affaire
                 avec gestion de messages d'erreur
Mots clefs ... :
*****************************************************************}
function TCALCULCOEF.CalculCoefFormule(pStAffaire, pStCodeFormule : String;
                                       pDtEcheance : TDateTime; pBoLog, pBoMsg : Boolean;
                                       pBoCoefRegul, pBoProchaine : Boolean;
                                       pTobParamFormule, pTobRevision : Tob; pBoRecupTob : Boolean): Boolean;
var
  vRdErr          : Integer;
  vStIndiceErreur : String;

begin

  fDtEcheance := pDtEcheance;
  fBoCoefRegul := pBoCoefRegul;
  fBoProchaine := pBoProchaine;
  if LoadFormule(pStAffaire, pStCodeFormule) then
    if LoadValeursIndices(vStIndiceErreur) then
      if LoadValeurInitIndices(vStIndiceErreur) then
      begin
        if not fBoFormuleOk then
        begin
          GestionErreur(pBoLog, pBoMsg, format(TexteMsg[3], [fStAffaire, fStCodeFormule, fStLibFormule]));
          result := false;
        end
        else if not fBoParamFormuleOK then
        begin
          GestionErreur(pBoLog, pBoMsg, format(TexteMsg[8], [fStAffaire, fStCodeFormule, fStLibFormule]));
          result := false;
        end
        else
        begin
          // test si c'est une formule d'actualisation et si actualisation nécessaire
          // si formule d'actualisation, pas de revision
          // sinon revision
          if (TestFormuleActualisation and TestActualisation) or
             (not TestFormuleActualisation) then
          begin
            if TestFormuleActualisation then fSLLog.Add(DateToStr(Now) + ' ' + TexteMsg[11]);
            if CalculCoef then
            begin
              if pBoCoefRegul then
                vRdErr := InsertCoefRegul(pDtEcheance)
              else
                vRdErr := InsertCoefficients(pBoRecupTob, pTobParamFormule, pTobRevision);
              if vRdErr = -1 then
              begin
                GestionErreur(pBoLog, pBoMsg, format(TexteMsg[13], [fStAffaire, fStCodeFormule, fStLibFormule]));
                result := false;
              end
              else if vRdErr = -2 then
              begin
                GestionErreur(pBoLog, pBoMsg, TexteMsg[2]);
                result := false;
              end
              else
                result := True;
            end
            else
            begin
              GestionErreur(pBoLog, pBoMsg, TexteMsg[14]);
              result := false;
            end;
          end
          // formule d'actualisation, mais pas d'actualisation
          else if TestFormuleActualisation then
          begin
            result := false;
            GestionErreur(pBoLog, pBoMsg, format(TexteMsg[12], [fStAffaire]));
          end
          else
            result := false;
        end;
      end
      else
      begin
        result := False;
        GestionErreur(pBoLog, pBoMsg, format(TexteMsg[5], [fStAffaire, vStIndiceErreur]));
      end
    else
    begin
      result := False;
      GestionErreur(pBoLog, pBoMsg, format(TexteMsg[6], [fStAffaire, vStIndiceErreur]));
    end
  else
  begin
    result := False;
    if pBoLog then
      fSLLog.Add(DateToStr(Now) + ' ' + format(TexteMsg[7], [fStCodeFormule, fStLibFormule, fStAffaire]))
    else
    begin
        PGIBoxAF (DateToStr(Now) + ' ' + format(TexteMsg[7] , [fStCodeFormule, fStLibFormule, fStAffaire]),'');
        fSLLog.Add(DateToStr(Now) + ' ' + format(TexteMsg[7], [fStCodeFormule, fStLibFormule, fStAffaire]));
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 29/07/2003
Modifié le ... :   /  /
Description .. : Retourne la liste des valeurs initiales d'un indice
                 on ne retourne pas pour l'instant l'indice en erreur
Mots clefs ... :
*****************************************************************}
function TCALCULCOEF.LectureValeursInit(pStAffaire, pStCodeFormule : String;
                                        var pArValInitIndice : Array of double): Boolean;
var
  i               : Integer;
  vStIndiceErreur : String;

begin
  result := False;
  if LoadFormule(pStAffaire, pStCodeFormule) then
    if LoadValeurInitIndices(vStIndiceErreur) then
    begin
      for i := 1 to 10 do
        pArValInitIndice[i] := fArValInitIndice[i];
      Result := True;
    end;
end;

procedure TCALCULCOEF.GestionErreur(pBoLog, pBoMsg : boolean; pTexteMsg : String);
begin
  if pBoLog then fSLLog.Add(pTexteMsg);
  if pBoMsg then PGIBoxAF (pTexteMsg ,'');
end;
 
{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 01/04/2003
Modifié le ... :
Description .. : Calcul du coefficient pour l'ensemble des formules d'une affaire
                 remarque : la date d'echeance n'est utilisée que pour le calcul
                 des coefficients de regularisation
                 la date d'écheance est en fait la date de la piece
                 la Tob des formules n'est utilisée que pour les factures
Mots clefs ... :
*****************************************************************}
function TCALCULCOEF.CalculCoefRegularisation(pStAffaire : String; pDtEcheance : TDateTime): Boolean;
var
  i     : Integer;
  vSt   : String;
  vQr   : TQuery;
  vTob  : Tob;

begin

  // remise a jour des paramformule
  VerifieCoherence(pStAffaire, True);

  result := true;
  // recherche des formules d'une affaire
  vSt := 'SELECT AFC_FORCODE FROM AFPARAMFORMULE ';
  vSt := vSt + ' WHERE AFC_AFFAIRE = "' + pStAffaire + '"';

  vQr := nil;
  vTob := TOB.Create('AFPARAMFORMULE', nil, -1);
  Try
    vQR := OpenSql(vSt, True,-1,'',true);
    if Not vQR.Eof then
    begin
      vTob.LoadDetailDB('AFPARAMFORMULE','','',vQR,False,True);
      DebutLog('', fSLLog);
      for i := 0 to vTob.Detail.Count -1 do
        begin
          result := CalculCoefFormule(pStAffaire, vTob.Detail[i].Getvalue('AFC_FORCODE'), pDtEcheance,
                                      True, // gestion d'un fichier log
                                      False, // affichage messages
                                      True, // regularisation
                                      False, // prochaine revision
                                      nil, // pas retour des tob
                                      nil,
                                      false);
          if not result then break;
        end;
        FinLog('', fSLLog);
    end;
  Finally
    if vQR <> nil then ferme(vQR);
    vTob.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 01/04/2003
Modifié le ... :
Description .. : Calcul du coefficient pour l'ensemble des formules d'une affaire
                 remarque : la Tob des formules n'est utilisée que pour les factures
Mots clefs ... :
*****************************************************************}
function TCALCULCOEF.CalculCoefFacturation(pStAffaire : String; pTobFormule, pTobParamFormule, pTobRevision : Tob; pBoRecupTob : Boolean): Boolean;
var
  i     : Integer;
  vSt   : String;
  vQr   : TQuery;
  vTob  : Tob;

begin

  // remise a jour des paramformule
  // création d'une ligne de parametrage avec les valeurs par defaut
  VerifieCoherence(pStAffaire, True);
  result := true;

  // automatique
  if not assigned(pTobFormule) then
  begin
    vSt := 'SELECT AFC_FORCODE FROM AFPARAMFORMULE ';
    vSt := vSt + ' WHERE AFC_AFFAIRE = "' + pStAffaire + '"';

    vQr := nil;
    vTob := TOB.Create('AFPARAMFORMULE', nil, -1);
    Try
      vQR := OpenSql(vSt, True,-1,'',true);
      if Not vQR.Eof then
      begin
         vTob.LoadDetailDB('AFPARAMFORMULE','','',vQR,False,True);
         DebutLog('', fSLLog);
         for i := 0 to vTob.Detail.Count -1 do
           begin
             result := CalculCoefFormule(pStAffaire, vTob.Detail[i].Getvalue('AFC_FORCODE'),
                                      -1,   // date d'echeance
                                      true, // fichier log uniquement
                                      true, // affichage messages                                      
                                      False, // regularisation
                                      True, // prochaine revision
                                      pTobParamFormule, // paramformule
                                      pTobRevision,     // nouveaux calculs
                                      pBoRecupTob);     // retour des données par tob

             if not result then break;
           end;
         FinLog('', fSLLog);
      end;
    Finally
      if vQR <> nil then ferme(vQR);
      vTob.Free;
    end;
  end
  else
  begin
    DebutLog('', fSLLog);
    for i := 0 to pTobFormule.Detail.Count -1 do
    begin
      if pTobFormule.Detail[i].GetValue('RECALCUL') = 'X' then
        if pTobFormule.Detail[i].GetValue('PROCHAINE') = 'X' then
        begin
          result := CalculCoefFormule(pStAffaire, pTobFormule.Detail[i].Getvalue('FORMULE'),
                                      -1,    // date d'echeance
                                      false, // gestion d'un fichier log
                                      True, // affichage messages
                                      False, // regularisation
                                      True,  // prochaine revision
                                      pTobRevision,     // nouveaux calculs
                                      pTobParamFormule, // paramformule
                                      pBoRecupTob);     // retour des données par tob
          if not result then break;
        end
        else
        begin
          result := CalculCoefFormule(pStAffaire, pTobFormule.Detail[i].Getvalue('FORMULE'),
                                      -1,    // date d'echeance
                                      false, // gestion d'un fichier log
                                      True, // affichage messages
                                      False, // regularisation
                                      False, // prochaine revision
                                      pTobRevision,     // nouveaux calculs
                                      pTobParamFormule, // paramformule
                                      pBoRecupTob);    // retour des données par tob
          if not result then break;
        end;
    end;
    FinLog('', fSLLog);
  end;
end;



{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 13/06/2003
Modifié le ... :   /  /
Description .. : application des coefficients d'une formule
Mots clefs ... :
*****************************************************************}
procedure TCALCULCOEF.AppliquerCoef(pBoAncienCoef, pBoMajDates : Boolean);
var
  vSt : String;

begin
 
  //ancien coef
  if pBoAncienCoef then
  begin
    vSt :='Update AFREVISION SET AFR_COEFAPPLIQUE=AFR_DERNIERCOEF, ' ;
    vSt := vSt +' AFR_OKCOEFAPPLIQUE="X" ' ;
    vSt := vSt +' WHERE AFR_AFFAIRE="'+ fStAffaire + '"';
    vSt := vSt +' AND AFR_FORCODE="'+ fStCodeFormule +'"' ;
    executesql(vSt);
  end
 
  // application du coef calculé
  else
  begin 
    vSt :='Update AFREVISION SET AFR_OKCOEFAPPLIQUE="X" ' ;
    vSt := vSt +' WHERE AFR_AFFAIRE="'+ fStAffaire + '"';
    vSt := vSt +' AND AFR_FORCODE="'+ fStCodeFormule +'"' ;
    executesql(vSt);
  end;

  // si le coef appliquer est 0, alors on le force à 1
  vSt :='Update AFREVISION SET AFR_COEFAPPLIQUE = 1 ' ;
  vSt := vSt +' WHERE AFR_AFFAIRE="'+ fStAffaire + '"';
  vSt := vSt +' AND AFR_FORCODE="'+ fStCodeFormule +'"' ;
  vSt := vSt +' AND AFR_COEFAPPLIQUE= 0' ;
  executesql(vSt);

  if pBoMajDates then
    MajDatesRevisionsApresApplication(True, False, nil);
         
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 01/08/2003
Modifié le ... :   /  /
Description .. : annuler l'application des coefficients
Mots clefs ... :
*****************************************************************}
function TCALCULCOEF.DesappliquerCoef(pDtDate : TDateTime) : Boolean;
var
  vSt     : String;
  vSLLog  : TStringList;
  vQr     : Tquery;

begin

  // si des coef ont ete calculés a une date ulterieure
  // on ne peut pas desappiquer les coefficients
  vSt := 'SELECT AFR_AFFAIRE FROM AFREVISION ';
  vSt := vSt + ' WHERE AFR_AFFAIRE="' + fStAffaire + '"';
  vSt := vSt + ' AND AFR_FORCODE="' + fStCodeFormule +'"' ;
  vSt := vSt + ' AND AFR_DATECALCCOEF > "' + usDateTime(pDtDate) +'"' ;

  vQr := nil;
  Try
    vQR := OpenSql(vSt, True,-1,'',true);
    if Not vQR.Eof then
    begin
      DebutLog('', vSLLog);         
      vSLLog.Add(format(TexteMsg[16], [fStAffaire, fStCodeFormule, dateToStr(pDtDate)]));
      result := false;
      FinLog('', vSLLog);
    end
    else
    begin
      // application du coef calculé ou de l'ancien coef
      vSt := 'Update AFREVISION SET AFR_OKCOEFAPPLIQUE= "-" ' ;
      vSt := vSt + ' WHERE AFR_AFFAIRE="' + fStAffaire + '"';
      vSt := vSt + ' AND AFR_FORCODE="' + fStCodeFormule +'"' ;
      vSt := vSt + ' AND AFR_DATECALCCOEF="' + usDateTime(pDtDate) +'"' ;
      executesql(vSt);

      if not MajDatesRevisionsApresDesapplication(fStAffaire, fStCodeFormule, pDtDate) then
      begin
        DebutLog('', vSLLog);
        vSLLog.Add(format(TexteMsg[15], [fStAffaire, fStCodeFormule, dateToStr(pDtDate)]));
        result := false;
        FinLog('', vSLLog);
      end
      else
        result := True;
    end;
  Finally
    Ferme(vQr);
  end;
end;

function TCALCULCOEF.CalculFormuleAuto(pTob : Tob) : Boolean;
var
  vIn             : TDerniereOuProchaine;
  vDt             : TDateTime;

begin
  DebutLog('', fSLLog);
  fBoCoefRegul := False;
  vIn := DerniereOuProchaine(pTob.GetValue('AFC_AFFAIRE'), pTob.Getvalue('AFC_FORCODE'));
  if vIn = Derniere then
  begin
    fBoProchaine  := False;
    fDtEcheance   := pTob.Getvalue('AFC_LASTDATECALC');
                                     
    result := CalculCoefFormule(pTob.GetValue('AFC_AFFAIRE'),
                                pTob.Getvalue('AFC_FORCODE'),
                                pTob.Getvalue('AFC_LASTDATECALC'),
                                True, // gestion d'un fichier log
                                False, // affichage messages
                                False, // regularisation
                                False, // prochaine revision
                                nil, // pas retour des tob
                                nil,
                                false);
  end
  else
  begin
    fBoProchaine := True;

    if pTob.getvalue('AFC_NEXTDATEAPP') = idate2099 then
      vDt := pTob.Getvalue('AFC_PREMIEREDATE')
    Else
      vDt := pTob.Getvalue('AFC_NEXTDATEAPP');
    fDtEcheance  := vDt;
 
    result := CalculCoefFormule(pTob.GetValue('AFC_AFFAIRE'),
                          pTob.Getvalue('AFC_FORCODE'),
                          vDt,
                          false, // gestion d'un fichier log
                          False, // affichage messages
                          False, // regularisation
                          True, // prochaine revision
                          nil, // pas retour des tob
                          nil,
                          false);
  end;
  FinLog('', fSLLog);
end;

Function TCALCULCOEF.DerniereOuProchaine(pStAffaire, pStFormule : String) : TDerniereOuProchaine;
Var
  St  : String;
  Q   : TQuery;
  T   : Tob;

begin

  T := TOB.Create('Derniere Ou Prochaine ', nil, -1);
  st:='SELECT MAX(AFR_DATECALCCOEF) AS MADATE, AFR_OKCOEFAPPLIQUE FROM AFREVISION  WHERE AFR_AFFAIRE = "'+ pStAffaire +'"' ;
  st:=st+' AND AFR_FORCODE = "'+ pStFormule +'" AND AFR_COEFREGUL = "-" GROUP BY AFR_OKCOEFAPPLIQUE ' ;

  Q:=Nil ;
  try
    Q := OpenSQL(St, TRUE,-1,'',true);
    T.LoadDetailDB('','','',Q,false) ;
  finally
    Ferme(Q) ;
  end ;

  if (T.Detail.Count > 1) and
     (T.detail[0].GetValue('MADATE') = T.detail[1].GetValue('MADATE')) then
  begin
    // on a les deux cas
    result:=Non ;
  end
  else
  begin
    if T.Detail.Count>0 then
    begin
      if T.Detail[0].getvalue('AFR_OKCOEFAPPLIQUE')='X' then
       result:=Prochaine
      else
       result := Derniere; // on peut recalculer
    end
    else
    begin
      result := Prochaine; // prochaine
    end ;
  end ;
  T.free ;
end ;


{procedure TCALCULCOEF.OnFinalization(Sender: TObject);
begin
;
end;
}

procedure TCALCULCOEF.OnFunction(Sender: TObject; FuncName: String;
  Params: array of variant; var Result: Variant);
var
  value : double;
  decimal : integer;
begin
if FuncName='@ARR' then
  begin
  value := Params[1];
  decimal := Params[2];
  Result := strToFloat(format('%.' + intToStr(decimal) + 'f', [value]));
  end;
end;

{procedure TCALCULCOEF.OnGetList(Sender: TObject; IdentList: String;
  var List: TStringList);
begin
;
end;
}

procedure TCALCULCOEF.OnGetVar(Sender: TObject; VarName: String;
  VarIndx: Integer; var Value: variant);
var
  i : Integer;
               
begin
  for i := 1 to fInNbindices do
  begin
    if ('IND'+ intToStr(i)) = VarName then
    begin                   
      case VarIndx of
        -1 : Value := fArValIndice[i];
        0 : Value := fArValInitIndice[i];
        end;
      break;
    end;
  end;
end;

{procedure TCALCULCOEF.OnInitialization(Sender: TObject);
begin
;
end;

procedure TCALCULCOEF.OnSetVar(Sender: TObject; VarName: String;
  VarIndx: Integer; Value: variant);
begin
;
end;
}

function TCALCULCOEF.FormuleEditionDetail(pStAffaire, pStCodeFormule : String; var pStFormule : String) : Boolean;
var
  i                   : Integer;
  vStIndiceErreur     : String;
//  PosVirgule          : Integer;
//  PosParentheseFermee : Integer;
//  StaJeter            : String;

begin             
  result := True;
  if LoadFormule(pStAffaire, pStCodeFormule) then
    if LoadValeursIndices(vStIndiceErreur) then
      if LoadValeurInitIndices(vStIndiceErreur) then
      begin
        pStFormule := fStFormule;
        for i := 1 to fInNbIndices do
        begin
          pStFormule:=Stringreplace(pStFormule, '[' + fArCodeIndice[i] + cStIndiceZero + ']', FloatToStr(fArValInitIndice[i]), [rfReplaceAll]);
          pStFormule:=Stringreplace(pStFormule, '[' + fArCodeIndice[i] + ']', FloatToStr(fArValIndice[i]), [rfReplaceAll]);
          pStFormule:=Stringreplace(pStFormule, ',' , '.', [rfReplaceAll]);
        end;
                  
        // on ne gere pas les arrondis
        result := not (pos(';',pStFormule)>0);
      end;
end;

end.
