{***********UNITE*************************************************
Auteur  ...... : NA
Créé le ...... : 19/10/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : INTEGPREPAIE () Intégration de la présence en paie
Mots clefs ... : TOF;INTEGPREPAIE
*****************************************************************}
Unit UTOFPGintegprespaie ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     HDB,
     fe_main,
{$else}
     eMul,
//     uTob,
     maineagl,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HTB97,
     P5Def,
     utof,
     Utob,
     Hqry,
     HStatus,
     ed_tools,
     AglInit;


Type
  TOF_PGINTEGPREPAIE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
 private
   {$IFNDEF EAGLCLIENT}
    Liste: THDBGrid;
{$ELSE}
    Liste: THGrid;
{$ENDIF}
    typetraitement, typeaffich: string;
    DateDeb, DateFin: THedit;
    TOB_CompteursPresence : TOB;
    Tob_Salaries : TOB;
    procedure AlimenteHISTOSAISRUB;
    procedure SupHISTOSAISRUB;
    procedure ChargecompteursPresence;
    procedure ChargeSalaries;
    procedure changeaffichage(sender : Tobject);
    procedure Lecturesalaries(var continuer : string);
    procedure Integrepaie(sender : Tobject);
    procedure Datedebexit(sender : Tobject);
    procedure Datefinexit(sender : Tobject);
    procedure SupIntegrepaie(sender : Tobject);
    function rechercheordre(salarie: string) : integer;
 end;

Implementation
uses PgOutils2;

procedure TOF_PGINTEGPREPAIE.OnNew ;
begin
  Inherited ;
  setcontroltext('CBAFFICHER', 'AIP');
end ;

procedure TOF_PGINTEGPREPAIE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGINTEGPREPAIE.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 19/10/2007
Modifié le ... :   /  /
Description .. : chargement
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGINTEGPREPAIE.OnLoad ;
begin
  Inherited ;
  // si traitement intégration en cours ; ne pas recharger la clause "WHERE"
  if typeaffich = 'AFF' then changeaffichage(nil);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 19/10/2007
Modifié le ... :   /  /    
Description .. : On argument
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGINTEGPREPAIE.OnArgument (S : String ) ;
var
Btn : TToolBarButton97;
DebPer, FinPer, ExerPerEncours : string;
OkOk: Boolean;
num : integer;
typeaffichage: THValComboBox;
begin
  Inherited ;
  setcontroltext('CBAFFICHER', 'AIP');
  typeaffich := 'AFF'; // en affichage du mul (traitement non en cours)
  typeaffichage:= THValComboBox (GetControl ('CBAFFICHER'));
  typeaffichage.OnChange := Changeaffichage;

  for Num := 1 to 4 do
  begin
  VisibiliteChampSalarie(IntToStr(Num), GetControl('PYP_TRAVAILN' + IntToStr(Num)), GetControl('TPYP_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PYP_CODESTAT'), GetControl('TPYP_CODESTAT'));

  Btn := TToolBarButton97(GetControl('Bouvrir'));
  if btn <> nil then Btn.Onclick := integrepaie;

  Btn := TToolBarButton97(GetControl('B_SUPPRIMER'));
  if btn <> nil then Btn.Onclick := SUPintegrepaie;

  DateDeb := ThEdit(getcontrol('PERIODEDEBUT'));
  if DateDeb <> nil then DateDeb.OnExit := DateDebExit;
  DateFin := ThEdit(getcontrol('PERIODEFIN'));
  if DateFin <> nil then DateFin.OnExit := DateFinExit;

  OkOk := RendPeriodeEnCours(ExerPerEncours, DebPer, FinPer);
  if OkOk then
  begin
    if DateDeb <> nil then DateDeb.text := DebPer;
    if DateFin <> nil then DateFin.text := FinPer;
  end;

end ;

procedure TOF_PGINTEGPREPAIE.OnClose ;
begin
  Inherited ;
  if assigned(Tob_compteurspresence) then FreeandNil(Tob_compteurspresence);
  if assigned(Tob_Salaries) then FreeandNil(Tob_Salaries);
end ;

procedure TOF_PGINTEGPREPAIE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGINTEGPREPAIE.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 22/10/2007
Modifié le ... :   /  /    
Description .. : Sortie de date début exercice
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGINTEGPREPAIE.DateDebExit(Sender: TObject);
begin
  if not IsValidDate(DateDeb.Text) then
  begin
    PGIBox('La date de début n''est pas valide', Ecran.Caption);
    DateDeb.SetFocus;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 22/10/2007
Modifié le ... :   /  /    
Description .. : Sortie de fin exercice
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGINTEGPREPAIE.DateFinExit(Sender: TObject);
var
  Date1, Date2: TDateTime;
begin
  if not IsValidDate(DateFin.Text) then
  begin
    PGIBox('La date de fin n''est pas valide', Ecran.Caption);
    DateFin.SetFocus;
    exit;
  end;
  Date2 := StrToDate(DateFin.Text);
  Date1 := StrToDate(DateDeb.Text);
  if Date1 > Date2 then
  begin
    PGIBox('La date de début est supérieure à la date de fin', Ecran.Caption);
    DateDeb.SetFocus;
    DateFin.Text := '';
    exit;
  end;
  //if (FINDEMOIS(Date1) <> FINDEMOIS(Date2)) then
  //begin
  //  PGIBox('Les dates de début et de fin doivent être comprise dans le même mois', Ecran.Caption);
  //  DateFin.SetFocus;
  //  exit;
  //end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 22/10/2007
Modifié le ... :   /  /    
Description .. : Changement du type d'affichage
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGINTEGPREPAIE.Changeaffichage (sender : Tobject);
var
Datepres : Tdatetime;
begin
  typeaffich := 'AFF'; // affichage du mul
  Datepres := StrToDate(getcontrolText('DATEDEBUT'));
  if GetControlText('CBAFFICHER')= 'AIP'  then // à intégrer en paie
  begin
    typetraitement := 'AIP';
    setcontrolvisible('B_SUPPRIMER', false);
    setcontrolvisible('Bouvrir', true);
    Ecran.Caption := 'Intégration des compteurs de présence en paie';
    updateCaption(Ecran);
    if datepres = iDate1900  then
    SetControlText('XX_WHERE','PYP_PGINDICATPRES <> "AIN" AND PYP_PGINDICATPRES <> "INP"'+
     ' AND PYR_INTEGREPAIE = "X" ')
    else
     SetControlText('XX_WHERE','PYP_PGINDICATPRES <> "AIN" AND PYP_PGINDICATPRES <> "INP" '+
     ' AND  PYR_INTEGREPAIE = "X" AND PYP_DATEDEBUTPRES >="'+UsDateTime(Datepres)+'" ');
  end
  else
   // compteurs intégrés en paie et non valorisés
  begin
    typetraitement := 'SUP';
    setcontrolvisible('B_SUPPRIMER', true);
    setcontrolvisible('Bouvrir', false);
    Ecran.Caption := 'Compteurs de présence intégrés en paie non valorisés';
    updateCaption(Ecran);
    if datepres <> iDate1900  then
    SetControlText('XX_WHERE','PYP_PGINDICATPRES = "AIN" ')
    else
     SetControlText('XX_WHERE','PYP_PGINDICATPRES = "AIN" AND PYP_DATEDEBUTPRES >="'+UsDateTime(Datepres)+'" ');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 19/10/2007
Modifié le ... :   /  /
Description .. : Lecture des salariés sélectionnés
Mots clefs ... :
*****************************************************************}
procedure TOF_PGINTEGPREPAIE.lecturesalaries(var continuer : string);
var
Suite, i : integer;
salarie , st, compteur : string;
Q_mul : THQuery;
datedebutpres, datefinpres : Tdatetime;

begin
{$IFNDEF EAGLCLIENT}
Liste:= THDBGrid(GetControl('FListe'));
{$ELSE}
Liste:= THGrid(GetControl('FListe'));
{$ENDIF}
  continuer := 'NON';
   if Liste <> nil then
   begin   // D1
     if (Liste.NbSelected = 0) and (not Liste.AllSelected) then
     begin
       PGIBOX('Aucune ligne n''est sélectionnée.', Ecran.Caption);
       exit;
     end;


// lecture de la liste des salariés
     Q_mul := THQuery(Ecran.FindComponent('Q'));
     St := '';
     if (Liste.AllSelected = false) then
           // lecture de chaque salarié sélectionné
      begin
         for i := 0 to Liste.NbSelected - 1 do
         begin
           {$IFDEF EAGLCLIENT}
            TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
           {$ENDIF}
            Liste.GotoLeBOOKMARK(i);
            salarie := Q_mul.findfield('PYP_SALARIE').asstring;
            datedebutpres := Q_Mul.findfield('PYP_DATEDEBUTPRES').asdatetime;
            datefinpres := Q_Mul.findfield('PYP_DATEFINPRES').asdatetime;
            compteur := Q_Mul.findfield('PYP_COMPTEURPRES').asstring;
            St := St + ' (PYP_SALARIE="' + Salarie + '" AND PYP_DATEDEBUTPRES = "'+Usdatetime(datedebutpres)+'"'+
             ' AND PYP_DATEFINPRES = "'+usdatetime(datefinpres)+'" AND PYP_COMPTEURPRES = "'+compteur+'")  OR';
            end;
       Liste.ClearSelected;
     end; 

    if St <> '' then
    begin
      St := Copy(St, 1, Length(st) - 2)  ;
      SetControlText('XX_WHERE',  St);
      typeaffich := 'TRT'; // traitement en cours
      TFMul(Ecran).BCherche.Click;
      setcontrolvisible('B_SUPPRIMER', false);
      setcontrolvisible('Bouvrir', false);
    end;
    if typetraitement  <> 'SUP' then
    suite := Pgiask('Intégration en paie des compteurs de présence sélectionnés. Voulez vous poursuivre?' ,Ecran.Caption)
    else
    suite := Pgiask('Suppression de l''Intégration en paie des compteurs de présence sélectionnés. Voulez vous poursuivre?' ,Ecran.Caption);

    if (suite <> mrYes) then
    begin
     changeaffichage(nil);
     TFMul(Ecran).BCherche.Click;
     exit;
    end;
    Continuer := 'OUI';
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 19/10/2007
Modifié le ... :   /  /
Description .. : Intégration des compteurs de présence en paie
Mots clefs ... :
*****************************************************************}
procedure TOF_PGINTEGPREPAIE.Integrepaie(sender : Tobject);
var
Continuer : string;
begin

   Lecturesalaries(continuer);  // Lecture des salariés sélectionnés

   if Continuer = 'OUI' then
   begin

     if Tob_CompteursPresence = nil then
     chargecompteurspresence; // charge les compteurs pour récupérer le type de champ de la rubrique à alimenter

     if Tob_Salaries = nil then
     chargeSalaries;      // charge les salariés

     AlimenteHISTOSAISRUB;  // alimente HISTOSAISRUB
   end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 22/10/2007
Modifié le ... :   /  /    
Description .. : Chargement des compteurs de présence
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGINTEGPREPAIE.Chargecompteurspresence;
var
st : string;
Q : Tquery;
begin
  Tob_CompteursPresence:= TOB.Create('Compteurs de présence', nil, -1);
  St := 'SELECT PYR_COMPTEURPRES, PYR_DATEVALIDITE, PYR_TYPECHAMPRUB FROM COMPTEURPRESENCE WHERE ##PYR_PREDEFINI##'+
  ' PYR_INTEGREPAIE="X" ';
  Q := OpenSql(st, TRUE);
  Tob_compteursPresence.LoadDetailDB('COMPTEURPRESENCE', '', '', Q, FALSE, FALSE);
  ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 23/10/2007
Modifié le ... :   /  /    
Description .. : Charge info confidentiel du salarié
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGINTEGPREPAIE.ChargeSalaries;
var
st : string;
Q : Tquery;
DD: Tdatetime;
begin
  DD := StrToDate(DateDeb.Text);
  Tob_Salaries:= TOB.Create('Salarie', nil, -1);
  St := 'SELECT PSA_SALARIE, PSA_CONFIDENTIEL FROM SALARIES WHERE PSA_DATESORTIE>="'+UsDateTime(DD)+'" OR'+
        ' PSA_DATESORTIE = "'+UsDateTime (Idate1900)+'" OR'+
        ' PSA_DATESORTIE IS NULL';
  Q := OpenSql(st, TRUE);
  Tob_Salaries.LoadDetailDB('SALARIES', '', '', Q, FALSE, FALSE);
  ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 22/10/2007
Modifié le ... :   /  /    
Description .. : Alimente la table HISTOSAISRUB
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGINTEGPREPAIE.AlimenteHISTOSAISRUB;
var
Tob_histosaisrub, T, t_cptpres, T_sal : TOB;
typealim, compteur , datevalidite, salarie, salariesauv, confidentiel, typecalpres  : string;
ordre : integer;
DD, DF, datedebutpres, datefinpres : Tdatetime;
Montant : double;
begin
 try
  BeginTrans;
  Tob_histosaisrub := TOB.Create('Compteurs intégrés dans saisie par rubrique', nil, -1);
  InitMoveProgressForm(nil, 'Intégration des compteurs de présence en paie', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
  InitMove(TFmul(Ecran).Q.RecordCount, '');
  salarie := '';
  salariesauv := '';
  DD := StrToDate(DateDeb.Text);
  DF := StrToDate(Datefin.Text);
  TFmul(Ecran).Q.First;
  while not TFmul(Ecran).Q.EOF do
  begin
     MoveCur(False);
     MoveCurProgressForm;
     T := TOB.Create('HISTOSAISRUB', Tob_HistoSaisRub, -1);
     if T <> nil then
     begin
        salarie := TFMul(Ecran).Q.Findfield('PYP_SALARIE').asstring;

        if salarie <> salariesauv then
        begin // si rupture sur le salarié
          ordre := rechercheordre(salarie);  // recherche ordre dans HISTOSAISRUB
          // recherche confidentiel
          T_sal := Tob_Salaries.findfirst(['PSA_SALARIE'],[salarie], true);
          confidentiel := T_sal.getvalue('PSA_CONFIDENTIEL');
          salariesauv := salarie;
        end
        else
        ordre := ordre + 1;

        typecalpres := TFMul(Ecran).Q.Findfield('PYP_TYPECALPRES').asstring;
        Datedebutpres := strtodate(TFMul(Ecran).Q.Findfield('PYP_DATEDEBUTPRES').asstring);
        Datefinpres := strtodate(TFMul(Ecran).Q.Findfield('PYP_DATEFINPRES').asstring);
        T.PutValue('PSD_ORIGINEMVT', 'PRE');
        T.PutValue('PSD_SALARIE', salarie);
        T.PutValue('PSD_ETABLISSEMENT',TFMul(Ecran).Q.Findfield('PYP_ETABLISSEMENT').asstring);
        T.PutValue('PSD_DATEDEBUT', DD);
        T.PutValue('PSD_DATEFIN', DF);
        T.PutValue('PSD_RUBRIQUE', TFMul(Ecran).Q.Findfield('PYR_RUBRIQUE').asstring);
        compteur := TFMul(Ecran).Q.Findfield('PYP_COMPTEURPRES').asstring;
        Datevalidite := TFMul(Ecran).Q.Findfield('DATEVALIDITE').asstring;
        T_cptpres := Tob_Compteurspresence.findfirst(['PYR_COMPTEURPRES', 'PYR_DATEVALIDITE'],[compteur,datevalidite], true);
        typealim := T_cptpres.getvalue('PYR_TYPECHAMPRUB');
        T.PutValue('PSD_TYPALIMPAIE', typealim);
        T.PutValue('PSD_ORDRE', ordre);
        T.PutValue('PSD_CONFIDENTIEL', confidentiel);
       //    T.PutValue('PSD_LIBELLE',TFMul(Ecran).Q.Findfield('PYR_LIBELLE').asstring);
        montant :=TFMul(Ecran).Q.Findfield('PYP_QUANTITEPRES').asfloat ;
        if typealim = 'BAS' then T.PutValue('PSD_BASE', montant) else
        if typealim = 'MON' then T.PutValue('PSD_MONTANT', montant) else
        if typealim = 'COE' then T.PutValue('PSD_COEFF', montant) else
        if typealim = 'TAU' then T.PutValue('PSD_TAUX', montant);

     // Mise à jour de la table PRESENCESALARIE
     ExecuteSQL('UPDATE PRESENCESALARIE SET PYP_PGINDICATPRES = "AIN" , PYP_DATEDEBUTBUL = "'+usdatetime(DD)+'"'+
                ', PYP_DATEFINBUL = "'+usdatetime(DF)+'" WHERE PYP_SALARIE = "'+salarie+'" AND '+
                'PYP_COMPTEURPRES = "'+compteur+'" AND PYP_TYPECALPRES = "'+typecalpres+'" AND '+
                'PYP_DATEDEBUTPRES = "'+ usdatetime(datedebutpres)+'" AND PYP_DATEFINPRES = "'+ usdatetime(datefinpres) +'"');

     end;
     TFmul(Ecran).Q.Next;
    end;
   Tob_histosaisrub.InsertDB(nil, false);
   PGIInfo('Traitement terminé#10#13L''intégration en paie des compteurs a été effectuée normalement.', Ecran.Caption);

   COMMITTRANS;
  except
   PGIError('Erreur lors du traitement#10#13L''intégration en paie des compteurs n''a pas été effectuée.', Ecran.Caption);
   ROLLBACK;
  end;
  FiniMove;
  FiniMoveProgressForm;
 if assigned(tob_histoSaisrub) then FreeandNil(tob_histosaisrub);
 changeaffichage(nil);
 TFMul(Ecran).BCherche.Click;
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 23/10/2007
Modifié le ... :   /  /    
Description .. : Recherche du n°ordre disponible pour le salarié dans 
Suite ........ : HISTOSAISRUB
Mots clefs ... : 
*****************************************************************}
function TOF_PGINTEGPREPAIE.rechercheordre(salarie: string) : integer;
var
st : string;
Q : TQUERY;
DD, DF : Tdatetime;

begin
result := 0;
DD := strtodate(datedeb.Text);
DF := strtodate(datefin.Text);

st := 'SELECT MAX(PSD_ORDRE)  AS ORDRE ' +
      'FROM HISTOSAISRUB ' +
      'WHERE PSD_ORIGINEMVT = "PRE" ' +
      'AND PSD_SALARIE="' + salarie + '" ' +
      'AND PSD_DATEDEBUT="' + usdatetime(DD) + '" ' +
      'AND PSD_DATEFIN="' + usdatetime(DF) + '" ';
    Q := OpenSql(st, TRUE);
    if not Q.EOF then result :=  Q.FindField('ORDRE').AsInteger + 1
    else result := 1;
    Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 25/10/2007
Modifié le ... :   /  /    
Description .. : Suppression de l'intégration de la présence en paie
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGINTEGPREPAIE.SupIntegrepaie (sender : Tobject);
var
continuer : string;
begin

  Lecturesalaries(Continuer); // lecture des salariés sélectionnés

  if Continuer = 'OUI' then
  SupHISTOSAISRUB; // Suppression des enregs dans HISTOSAISBUL et mise à jour PRESENCESALARIE

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 25/10/2007
Modifié le ... :   /  /    
Description .. : Traitement de suppression de la présence en paie : 
Suite ........ : suppression dans HISTOSAISRUB et mise à jour de 
Suite ........ : PRESENCESALARIE
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGINTEGPREPAIE.SupHISTOSAISRUB;
var
compteur , datevalidite, salarie, etatpres, typecalpres, indicatpres,st, rubrique: string;
DD, DF, datedebutpres, datefinpres : Tdatetime;
Montant : double;
Q : TQUERY;

begin
 try
  BeginTrans;
  InitMoveProgressForm(nil, 'Suppression de l''intégration en paie des compteurs de présence', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
  InitMove(TFmul(Ecran).Q.RecordCount, '');
  salarie := '';
  DD := StrToDate(DateDeb.Text);
  DF := StrToDate(Datefin.Text);
  TFmul(Ecran).Q.First;
  while not TFmul(Ecran).Q.EOF do
  begin
     MoveCur(False);
     MoveCurProgressForm;
     salarie := TFMul(Ecran).Q.Findfield('PYP_SALARIE').asstring;
     Datedebutpres := strtodate(TFMul(Ecran).Q.Findfield('PYP_DATEDEBUTPRES').asstring);
     Datefinpres := strtodate(TFMul(Ecran).Q.Findfield('PYP_DATEFINPRES').asstring);
     compteur := TFMul(Ecran).Q.Findfield('PYP_COMPTEURPRES').asstring;
     Datevalidite := TFMul(Ecran).Q.Findfield('DATEVALIDITE').asstring;
     montant :=TFMul(Ecran).Q.Findfield('PYP_QUANTITEPRES').asfloat ;
     etatpres:= TFMul(Ecran).Q.Findfield('PYP_ETATPRES').asstring;
     Typecalpres:= TFMul(Ecran).Q.Findfield('PYP_TYPECALPRES').asstring;
     Rubrique:= TFMul(Ecran).Q.Findfield('PYR_RUBRIQUE').asstring;

     // Mise à jour de la table PRESENCESALARIE

     // si compteur saisi : PGINDICATPRES = 'NON' sinon = 'CAL'
     if etatpres = 'SAI' then indicatpres := 'NON' else indicatpres := 'CAL';

     // recherche s'il existe une régularisation sur ce compteur
     st :='SELECT PYP_QUANTITEPRES  FROM PRESENCESALARIE WHERE  PYP_SALARIE = "'+salarie+'" AND'+
          ' PYP_COMPTEURPRES = "'+compteur+'" AND PYP_DATEDEBUTPRES = "'+ usdatetime(datedebutpres)+'" AND'+
          ' PYP_DATEFINPRES = "'+ usdatetime(datefinpres) +'" AND PYP_TYPECALPRES = "002"';
     q:= opensql(st, true);
     if NOT Q.EOF then
     begin
       // le mopntant du ccmpteur régularisé = montant du compteur calculé + montant régularisé
       // et suppression de la ligne régularisée
       montant := montant + Q.findfield('PYP_QUANTITEPRES').asfloat;
       Executesql('DELETE PRESENCESALARIE WHERE PYP_SALARIE = "'+salarie+'" AND'+
                ' PYP_COMPTEURPRES = "'+compteur+'" AND PYP_DATEDEBUTPRES = "'+ usdatetime(datedebutpres)+'" AND'+
                ' PYP_DATEFINPRES = "'+ usdatetime(datefinpres) +'" AND PYP_TYPECALPRES = "002"');
     end;


     ExecuteSQL('UPDATE PRESENCESALARIE SET PYP_PGINDICATPRES = "'+indicatpres+'" , PYP_DATEDEBUTBUL ='+
                ' "'+usdatetime(idate1900)+'", PYP_DATEFINBUL = "'+usdatetime(idate1900)+'" , PYP_QUANTITEPRES'+
                ' = '+STRFPOINT(montant)+' WHERE PYP_SALARIE = "'+salarie+'" AND PYP_COMPTEURPRES = "'+compteur+'" AND '+
                'PYP_DATEDEBUTPRES = "'+ usdatetime(datedebutpres)+'" AND PYP_DATEFINPRES = '+
                ' "'+ usdatetime(datefinpres) +'" AND PYP_TYPECALPRES = "'+typecalpres+'"');

     // Suppression de l'enregistrement dans la table HISTOSAISRUB
     ExecuteSql('DELETE HISTOSAISRUB WHERE PSD_ORIGINEMVT = "PRE" AND PSD_SALARIE = "'+salarie+'" AND'+
                ' PSD_DATEDEBUT = "'+ usdatetime(DD)+'" AND PSD_DATEFIN = "'+usdatetime(DF)+'" AND'+
                ' PSD_RUBRIQUE = "'+rubrique+'"');

     TFmul(Ecran).Q.Next;
    end;
    PGIINFO('Traitement terminé#10#13La suppression de l''intégration en paie des compteurs de présence a été effectuée normalement.', Ecran.Caption);

   COMMITTRANS;
  except
   PGIError('Erreur lors du traitement#10#13La suppression de l''intégration en paie des compteurs n''a pas été effectuée.', Ecran.Caption);
   ROLLBACK;
  end;
  FiniMove;
  FiniMoveProgressForm;

 changeaffichage(nil);
 TFMul(Ecran).BCherche.Click;
end;

Initialization
  registerclasses ( [ TOF_PGINTEGPREPAIE ] ) ;
end.
