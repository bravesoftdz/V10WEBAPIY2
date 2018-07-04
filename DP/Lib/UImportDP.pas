
unit UImportDP;

interface

uses Classes,
     uTOB,
{$IFDEF EAGLCLIENT}
{$ELSE}
    {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     dialogs;

type
  TtobArray = Array of TOB;


  TImportDP = class
  private
    listeFic_ : TStrings;//liste des fichiers générés
    listeIdent_ : TStrings;//liste des identifiants rencontrés
    listeChps_ : TStrings;//liste des champs d'une table
    listeTypes_ : TStrings;//type associé à chaque champ
    tabOffset_ : array of integer;//tableau des offset pour chaque champ
    nomFic_ : string;//chemin du fichier
    fichier_ : TextFile;//fichier
    Tobs_ : TtobArray;

  //Méthodes privées
    function __initExtract(ident : string) : string;
    function __ligneToTob( ligne, nomTable : string; idx : integer) : integer;
    function __ajusterTaille(s :string) : string;
    function __SplitFile : integer;
    function __ConstruireTOB (subfic, ident : string; index : integer) : integer;
    procedure __setListes(nom, typ : string; offset : integer);
    function __convertValue(val, typ : string) : variant;
  public
    //Méthodes publiques
    constructor Create;
    destructor Free;
    function LanceImport(fic : string) : integer;
  //POUR TESTS EXPORT
    function generateFile(nomTable : string) : integer;
    function writeligne(rsql : TQuery; var f : textfile; table : string): integer;
    function correctlength(s :string; l : integer) : string;
  end;

{-----------------IMPLEMENTATION-----------------------------------------------}
implementation

uses Sysutils,
     HEnt1,
     Hctrls,
     UtilTrans,
     HMsgBox;

{supervision du traitement global de l'import}
function TImportDP.LanceImport(fic : string) : integer;
var
  i,vrt : integer;
  chk : boolean;
begin
  if Not FileExists(fic) then//on quitte si le fichier n'existe pas
    begin
    result := -1;
    exit;
    end;
  result := 0;//fct optimiste !
  nomFic_ := fic;
  AssignFile(fichier_, nomFic_);
//splitter en petits fichiers
  if __SplitFile <> 0 then
    begin
    result := -1;
    exit;
    end;
  Setlength(Tobs_, listeIdent_.Count);//allocation du tableau.
//pour chaque fichier : construire TOB
  For i:= 0 to listeFic_.Count-1 do
    begin
    vrt := __construireTOB(listeFic_[i], listeIdent_[i], i);//remplit la TOB à partir du fichier
    if vrt <> 0 then result := -1;
    end;
//insérer toutes les TOBs
  Try
    BeginTrans;
    if result = -1 then Raise Exception.Create('Erreur d''import du TRA.');
    For i:= 0 to listeFic_.Count -1 do
      begin
      chk := Tobs_[i].InsertOrUpdateDB;//insertion dans la base
      if Not chk then
        Raise Exception.Create('Erreur d''insertion de la TOB dans la base.');//contrôle sur l'insertion des TOBs
      deleteFile(listeFic_[i]);//Supprimer fichiers temporaires
      end;
    CommitTrans;
  Except
    RollBack;
    result := -1;
  End;
end;
{-----------------------------------------------------------------------------}
{Sépare le fichier TRA en plusieurs petits regroupant les lignes par leur identifiant}
function TImportDP.__SplitFile : integer;
var
  subFic : TextFile;
  ligne, nomSubFic : string;
  tmp : string;
begin
  try
  reset(fichier_);
  While Not EOF(fichier_) do//parcours du fichier
    begin
    readln(fichier_, ligne);//lit chaque ligne
    tmp := __ajusterTaille(copy(ligne, 4, 3));//récupère l'identifiant
    if listeIdent_.IndexOf(tmp) = -1 then//si absent de la liste
      begin
      if listeIdent_.Count <> 0 then//ce n'est pas la 1ère ligne du fichier
        closeFile(subFic);//close car un autre est déjà ouvert
      listeIdent_.Add(tmp);//ajout à la liste des identifiants
      nomSubFic := tempFileName;//génère un fichier temporaire
      listeFic_.Add(nomSubFic);//l'ajoute à la liste des fichiers
      assignFile(subFic, nomSubFic);
      rewrite(subFic);
      end;
    //écrire ligne dans le fichier
    writeln(subFic, ligne);
    end;
  closeFile(subFic);//fermer le dernier sous-fichier
  closeFile(fichier_);
  result := 0;
  except
  result := -1;
  end;
end;
{-----------------------------------------------------------------------------}
{Construit une TOB à partir du fichier subfic}
function TImportDP.__ConstruireTOB (subfic, ident : string; index : integer) : integer;
var
  fic : TextFile;
  ligne, nomTable : string;
  chk : integer;
begin
  result := 0;
  listeChps_ := TStringList.Create;
  listeTypes_ := TStringList.Create;
  nomTable := __initExtract(ident);//initalise les listes

  Try
    Tobs_[index] := TOB.Create('TOB '+nomTable,nil, -1);//allocation TOB mere (virtuelle)  de la table
    assignFile(fic, subfic);
    reset(fic);
    While not EOF(fic) do //lecture ligne par ligne du fichier
      begin
      readln(fic, ligne);
      //extraction des données
      chk := __ligneToTob(ligne, nomTable, index);
      if chk <> 0 then result := -1;
      end;
    closeFile(fic);
  Finally
    //libération
    listeChps_.Free;
    listeTypes_.Free;
    Finalize(TabOffset_);
  end;
end;
{-----------------------------------------------------------------------------}
{Prend en paramètre une ligne du fichier et met les données dans la tob}
function TImportDP.__ligneToTob( ligne, nomTable : string; idx : integer) : integer;
var
  tobF : TOB;
  i,c : integer;
  val : string;
  donnee : variant;
begin
  result := 0;
  tobF := TOB.Create(nomTable, Tobs_[idx], -1);
  c := listeChps_.Count-1;
  For i := 0 to c do
    begin
    //extrait donnée de la ligne et enlève les espaces à gauche ou à droite
    val := __ajusterTaille(copy(ligne, TabOffset_[i], TabOffset_[i+1] - TabOffset_[i]));
    if (val <> '') And (val<> '@') then//on n'insère pas les chamsp vides
      begin
      try
        //conversion de la donnée dans le type voulu
        donnee := __convertValue(val, ListeTypes_[i]);
      except
        on E: Exception do
          begin
          pgiinfo('ERREUR CONVERSION ! '+E.Message);
          result := -1;
          end;
      end;//except
      try
        tobF.PutValue(listeChps_[i], donnee);//insertion dans la tob
      except
        on E: Exception do
          begin
          pgiinfo('ERREUR INSERTION ! '+E.Message);
          result := -1;
          end;
      end;//except
      end;
    end;//FOR
end;
{------------------------------------------------------------------------------}
{Renvoie la valeur convertie dans le type voulu}
// #### TODO: quel traitement pour les valeurs vides ? => pas transmises
function TImportDP.__convertValue(val, typ : string) : variant;
var
  d : TDateTime;
  n : double;
  i,dd,mm,yyyy : integer;
  t : string;
begin
  t := LowerCase(typ[1]);
  if (t = 'a') Or (t = 'b') then //chaîne
    result := val
  else if t = 'd' then//date
    begin
    try
      //extraction des valeurs des jour, mois, année
      dd := strToInt(copy(val,1,2));
      mm := strToInt(copy(val, 3,2));
      yyyy := strToInt(copy(val,5,4));
      d := EncodeDate(yyyy, mm, dd);//création de la date
      result := d;
    except
       On EConvertError Do result := EncodeDate(1945,12,31);//échec => date au 01/01/1900 par défaut;
    end;//except
    end
  else if typ = 'n10' then//entier
    begin
    try
      i := StrToInt(val);
      result := i;
    except
      On EConvertError Do result := 0;//échec
    end;
    end
  else if (typ = 'n20.2') or (typ = 'n20.9') then//numeric
    begin
    try
      n := StrToFloat(val);
      result := n;
    except
      On EConvertError Do result := 0.0;
    end;
    end
end;
{-----------------------------------------------------------------------------}
{Renvoie le nom de la table associée à l'identifiant passé en paramètre
 & remplit les listes des champs}
function TImportDP.__initExtract(ident : string) : string;
begin
  listeChps_.Clear;
  listeTypes_.Clear;
  if ident = 'T' then
    begin
    result := 'TIERS';
    SetLength(tabOffset_, 152);//dimensionne le tableau
    __setListes('T_AUXILIAIRE','a17', 7);
    __setListes('T_NATUREAUXI','a3', 24);
    __setListes('T_LIBELLE','a35', 27);
    __setListes('T_EAN','a17', 62);
    __setListes('T_COLLECTIF','a17', 79);
    __setListes('T_ABREGE','a17', 96);
    __setListes('T_COMMENTAIRE','a35', 113);
    __setListes('T_ADRESSE1','a35', 148);
    __setListes('T_ADRESSE2','a35', 183);
    __setListes('T_ADRESSE3','a35', 218);
    __setListes('T_CODEPOSTAL','a9', 253);
    __setListes('T_VILLE','a35', 262);
    __setListes('T_DIVTERRIT','a9', 297);
    __setListes('T_PAYS','a3', 306);
    __setListes('T_LANGUE','a3', 309);
    __setListes('T_DEVISE','a3', 312);
    __setListes('T_MULTIDEVISE','b1', 315);
    __setListes('T_NIF','a17', 316);
    __setListes('T_SIRET','a17', 333);
    __setListes('T_APE','a5', 350);
    __setListes('T_TELEPHONE','a25', 355);
    __setListes('T_FAX','a25', 380);
    __setListes('T_TELEX','a25', 405);
    __setListes('T_TELEPHONE2','a25', 430);
    __setListes('T_FACTURE','a17', 455);
    __setListes('T_PAYEUR','a17', 472);
    __setListes('T_APPORTEUR','a17', 489);
    __setListes('T_TRANSPORTEUR','a17', 506);
    __setListes('T_COEFCOMMA','n20.2', 523);
    __setListes('T_SECTEUR','a3', 543);
    __setListes('T_ZONECOM','a3', 546);
    __setListes('T_REPRESENTANT','a17', 549);
    __setListes('T_TARIFTIERS','a3', 566);
    __setListes('T_REMISE','n20.2', 569);
    __setListes('T_FRANCO','n20.2', 589);
    __setListes('T_ESCOMPTE','n20.2', 609);
    __setListes('T_QUALIFESCOMPTE','a3', 629);
    __setListes('T_MODEREGLE','a3', 632);
    __setListes('T_JOURPAIEMENT1','n10', 635);
    __setListes('T_JOURPAIEMENT2','n10', 645);
    __setListes('T_FACTUREHT','b1', 655);
    __setListes('T_REGIMETVA','a3', 656);
    __setListes('T_SOUMISTPF','b1', 659);
    __setListes('T_LETTRABLE','b1', 660);
    __setListes('T_CORRESP1','a17', 661);
    __setListes('T_CORRESP2','a17', 678);
    __setListes('T_TOTALDEBIT','n20.2', 695);
    __setListes('T_TOTALCREDIT','n20.2', 715);
    __setListes('T_DATECREATION','d8', 735);
    __setListes('T_DATEMODIF','d8', 743);
    __setListes('T_DATEOUVERTURE','d8', 751);
    __setListes('T_DATEFERMETURE','d8', 759);
    __setListes('T_FERME','b1', 767);
    __setListes('T_DATEDERNMVT','d8', 768);
    __setListes('T_DEBITDERNMVT','n20.2', 776);
    __setListes('T_CREDITDERNMVT','n20.2', 796);
    __setListes('T_NUMDERNMVT','n10', 816);
    __setListes('T_LIGNEDERNMVT','n10', 826);
    __setListes('T_CONFIDENTIEL','b1', 836);
    __setListes('T_SOLDEPROGRESSIF','b1', 837);
    __setListes('T_SAUTPAGE','b1', 838);
    __setListes('T_TOTAUXMENSUELS','b1', 839);
    __setListes('T_COUTHORAIRE','n20.2', 840);
    __setListes('T_SOCIETE','a3', 860);
    __setListes('T_RELANCEREGLEMENT','a3', 863);
    __setListes('T_RELANCETRAITE','a3', 866);
    __setListes('T_RESIDENTETRANGER','a3', 869);
    __setListes('T_NATUREECONOMIQUE','a3', 872);
    __setListes('T_MOTIFVIREMENT','a3', 875);
    __setListes('T_LETTREPAIEMENT','a3', 878);
    __setListes('T_RELEVEFACTURE','b1', 881);
    __setListes('T_FREQRELEVE','a3', 882);
    __setListes('T_JOURRELEVE','n10', 885);
    __setListes('T_UTILISATEUR','a3', 895);
    __setListes('T_RVA','a250', 898);
    __setListes('T_EMAIL','a250', 1148);
    __setListes('T_DERNLETTRAGE','a4', 1398);
    __setListes('T_JURIDIQUE','a3', 1402);
    __setListes('T_FORMEJURIDIQUE','a3', 1405);
    __setListes('T_PROFIL','a3', 1408);
    __setListes('T_TVAENCAISSEMENT','a3', 1411);
    __setListes('T_TOTDEBP','n20.2', 1414);
    __setListes('T_TOTCREP','n20.2', 1434);
    __setListes('T_TOTDEBE','n20.2', 1454);
    __setListes('T_TOTCREE','n20.2', 1474);
    __setListes('T_TOTDEBS','n20.2', 1494);
    __setListes('T_TOTCRES','n20.2', 1514);
    __setListes('T_TOTDEBANO','n20.2', 1534);
    __setListes('T_TOTCREANO','n20.2', 1554);
    __setListes('T_TOTDEBANON1','n20.2', 1574);
    __setListes('T_TOTCREANON1','n20.2', 1594);
    __setListes('T_DATEDERNRELEVE','d8', 1614);
    __setListes('T_SCORERELANCE','n10', 1622);
    __setListes('T_CREERPAR','a3', 1632);
    __setListes('T_EXPORTE','a3', 1635);
    __setListes('T_SCORECLIENT','n10', 1638);
    __setListes('T_PAYEURECLATEMENT','b1', 1648);
    __setListes('T_DATEDERNPIECE','d8', 1649);
    __setListes('T_NUMDERNPIECE','n10', 1657);
    __setListes('T_TOTDERNPIECE','n20.2', 1667);
    __setListes('T_TABLE0','a17', 1687);
    __setListes('T_TABLE1','a17', 1704);
    __setListes('T_TABLE2','a17', 1721);
    __setListes('T_TABLE3','a17', 1738);
    __setListes('T_TABLE4','a17', 1755);
    __setListes('T_TABLE5','a17', 1772);
    __setListes('T_TABLE6','a17', 1789);
    __setListes('T_TABLE7','a17', 1806);
    __setListes('T_TABLE8','a17', 1823);
    __setListes('T_TABLE9','a17', 1840);
    __setListes('T_CONSO','a3', 1857);
    __setListes('T_CREDITDEMANDE','n20.2', 1860);
    __setListes('T_CREDITACCORDE','n20.2', 1880);
    __setListes('T_DOSSIERCREDIT','a35', 1900);
    __setListes('T_DATECREDITDEB','d8', 1935);
    __setListes('T_DATECREDITFIN','d8', 1943);
    __setListes('T_CREDITPLAFOND','n20.2', 1951);
    __setListes('T_DATEPLAFONDDEB','d8', 1971);
    __setListes('T_DATEPLAFONDFIN','d8', 1979);
    __setListes('T_NIVEAURISQUE','a3', 1987);
    __setListes('T_AVOIRRBT','b1', 1990);
    __setListes('T_ISPAYEUR','b1', 1991);
    __setListes('T_CODEIMPORT','a35', 1992);
    __setListes('T_TIERS','a17', 2027);
    __setListes('T_NATIONALITE','a3', 2044);
    __setListes('T_PRENOM','a35', 2047);
    __setListes('T_JOURNAISSANCE','n10', 2082);
    __setListes('T_MOISNAISSANCE','n10', 2092);
    __setListes('T_ANNEENAISSANCE','n10', 2102);
    __setListes('T_COMPTATIERS','a3', 2112);
    __setListes('T_DEBRAYEPAYEUR','b1', 2115);
    __setListes('T_MOISCLOTURE','n10', 2116);
    __setListes('T_EURODEFAUT','b1', 2126);
    __setListes('T_PARTICULIER','b1', 2127);
    __setListes('T_SEXE','a3', 2128);
    __setListes('T_PASSWINTERNET','a20', 2131);
    __setListes('T_PRESCRIPTEUR','a17', 2151);
    __setListes('T_PUBLIPOSTAGE','b1', 2168);
    __setListes('T_ORIGINETIERS','a3', 2169);
    __setListes('T_SOCIETEGROUPE','a17', 2172);
    __setListes('T_PHONETIQUE','a35', 2189);
    __setListes('T_DATEPROCLI','d8', 2224);
    __setListes('T_ETATRISQUE','a3', 2232);
    __setListes('T_DOMAINE','a3', 2235);
    __setListes('T_DATEINTEGR','d8', 2238);
    __setListes('T_NIVEAUIMPORTANCE','a3', 2246);
    __setListes('T_CLETELEPHONE','a25', 2249);
    __setListes('T_ENSEIGNE','a70', 2274);
    __setListes('T_EMAILING','b1', 2344);
    __setListes('T_REGION','a9', 2345);
    __setListes('T_DELAIMOYEN','n10', 2354);
    tabOffset_[151] :=  2364;    
    end;
{  else if ident = 'YTC' then
  else if ident = 'ANN' then
  else if ident = 'ANB' then
  else if ident = 'ANL' then
  else if ident = 'US'  then
  else if ident = 'C'   then
  else if ident = 'DCI' then
  else if ident = 'DCT' then
  else if ident = 'DCV' then
  else if ident = 'DPD' then
  else if ident = 'DPE' then
  else if ident = 'DFI' then
  else if ident = 'DIS' then
  else if ident = 'DPM' then
  else if ident = 'DPO' then
  else if ident = 'DPT' then
  else if ident = 'DOR' then
  else if ident = 'DPA' then
  else if ident = 'DPP' then
  else if ident = 'DSO' then
  else if ident = 'DSC' then
  else if ident = 'DTC' then
  else if ident = 'DT1' then
  else if ident = 'DTP' then
  else if ((ident = 'CC') Or (ident = 'YX')) then  }

end;
{-----------------------------------------------------------------------------}
{enlève les espaces situées avant et après la chaîne}
function TImportDP.__ajusterTaille(s :string) : string;
var
  tmp : string;
  i :integer;
begin
  i := 0;
  while i< length(s) do
    begin
    inc(i);
    if s[i]<> ' ' then break;
    end;
  tmp := copy(s, i, length(s)-i+1);
  i := length(tmp)+1;
  while i>1 do
    begin
    dec(i);
    if tmp[i]<>' ' then break;
    end;
  tmp := copy(tmp,1, i);
  if tmp = ' ' then result := ''
  else result := tmp;
end;
{-----------------------------------------------------------------------------}
{ajoute le nom du champ, son type et son offset dans la liste correspondante}
procedure TImportDP.__setListes(nom, typ : string; offset : integer);
begin
  ListeChps_.Add(nom);
  ListeTypes_.Add(typ);
  TabOffset_[ListeChps_.Count-1] := offset;
end;
{-----------------------------------------------------------------------------}
constructor TImportDP.Create;
begin
  //creation des listes
  listeFic_ := TStringList.Create;
  listeIdent_ := TStringList.Create;
  listeFic_.Clear;
  listeIdent_.Clear;
//  SetLength(Tobs, 26);//allocation du tableau. -> est faite plus tard
end;
{-----------------------------------------------------------------------------}
destructor TImportDP.Free;
var i : integer;
begin
  for i := 0 to listeIdent_.Count-1 do
    Tobs_[i].Free;//libère les tobs
  //libération des listes
  listeFic_.Free;
  listeIdent_.Free;
  //libère les tableaux dynamiques
  Finalize(Tobs_);
end;
{-----------------------------------------------------------------------------}

{******************************************************************************
 *                 FONCTIONS D'EXPORT POUR TESTS                              *
 ******************************************************************************}

 {Ajuste la chaîne à la longueur désirée}
function TImportDP.correctlength(s :string; l : integer) : string;
begin
  if length(s) > l then//chaîne trop longue : on coupe
    result := copy(s,1 , l)
  else if length(s) < l then//chaîne trop courte : on bourre à droite
    begin
    if V_PGI.SAV And (s = '') then
      result := s+'@'+format_string(' ', l-length(s)-1)//version de debugage : fichier inexploitable !
    else
      result := s+format_string(' ', l-length(s))//version normale;
    end
  else if length(s) = l then//bonne longueur
    result := s;
end;
{------------------------------------------------------------------------------}
{écrit une ligne du TRA}
function TImportDP.writeligne( rsql : tquery; var f : textfile; table : string) : integer;
var
  s : string;
  formatDate : string;
begin
  result := 0;
  s := '';
  formatDate := ShortDateFormat;
  ShortDateFormat := 'ddmmyyyy';
  if table = 'TIERS' then
    begin
    s := '***T  ';//fixe+identifiant
    //remplir
    s := s+correctlength(rsql.findfield('T_AUXILIAIRE').asstring, 17);
    s := s+correctlength(rsql.findfield('T_NATUREAUXI').AsString, 3);
    s := s+correctlength(rsql.findfield('T_LIBELLE').AsString, 35);
    s := s+correctlength(rsql.findfield('T_EAN').AsString, 17);
    s := s+correctlength(rsql.findfield('T_COLLECTIF').AsString, 17);
    s := s+correctlength(rsql.findfield('T_ABREGE').AsString, 17);
    s := s+correctlength(rsql.findfield('T_COMMENTAIRE').AsString, 35);
    s := s+correctlength(rsql.findfield('T_ADRESSE1').Asstring, 35);
    s := s+correctlength(rsql.findfield('T_ADRESSE2').Asstring, 35);
    s := s+correctlength(rsql.findfield('T_ADRESSE3').asstring, 35);
    s := s+correctlength(rsql.findfield('T_CODEPOSTAL').asstring, 9);
    s := s+correctlength(rsql.findfield('T_VILLE').asstring, 35);
    s := s+correctlength(rsql.findfield('T_DIVTERRIT').asstring, 9);
    s := s+correctlength(rsql.findfield('T_PAYS').AsString, 3);
    s := s+correctlength(rsql.findfield('T_LANGUE').asstring, 3);
    s := s+correctlength(rsql.findfield('T_DEVISE').asstring, 3);
    s := s+correctlength(rsql.findfield('T_MULTIDEVISE').asstring, 1);
    s := s+correctlength(rsql.findfield('T_NIF').asstring, 17);
    s := s+correctlength(rsql.findfield('T_SIRET').asstring, 17);
    s := s+correctlength(rsql.findfield('T_APE').asstring, 5);
    s := s+correctlength(rsql.findfield('T_TELEPHONE').asstring, 25);
    s := s+correctlength(rsql.findfield('T_FAX').asstring, 25);
    s := s+correctlength(rsql.findfield('T_TELEX').asstring, 25);
    s := s+correctlength(rsql.findfield('T_TELEPHONE2').asstring, 25);
    s := s+correctlength(rsql.findfield('T_FACTURE').asstring, 17);
    s := s+correctlength(rsql.findfield('T_PAYEUR').asstring, 17);
    s := s+correctlength(rsql.findfield('T_APPORTEUR').asstring, 17);
    s := s+correctlength(rsql.findfield('T_TRANSPORTEUR').asstring, 17);
    s := s+alignDroite(StrfMontant(rsql.findField('T_COEFCOMMA').AsFloat,20,2,'',false), 20);
    s := s+correctlength(rsql.findfield('T_SECTEUR').asstring, 3);
    s := s+correctlength(rsql.findfield('T_ZONECOM').asstring, 3);
    s := s+correctlength(rsql.findfield('T_REPRESENTANT').asstring, 17);
    s := s+correctlength(rsql.findfield('T_TARIFTIERS').asstring,3);
    s := s+alignDroite(StrfMontant(rsql.findField('T_REMISE').AsFloat,20,2,'',false), 20);
    s := s+alignDroite(StrfMontant(rsql.findField('T_FRANCO').AsFloat,20,2,'',false), 20);
    s := s+alignDroite(StrfMontant(rsql.findField('T_ESCOMPTE').AsFloat,20,2,'',false), 20);
    s := s+correctlength(rsql.findfield('T_QUALIFESCOMPTE').AsString, 3);
    s := s+correctlength(rsql.findfield('T_MODEREGLE').AsString, 3);
    s := s+alignDroite(StrfMontant(rsql.findField('T_JOURPAIEMENT1').AsFloat,10,0,'',false), 10);
    s := s+alignDroite(StrfMontant(rsql.findField('T_JOURPAIEMENT2').AsFloat,10,0,'',false), 10);
    s := s+correctlength(rsql.findfield('T_FACTUREHT').AsString, 1);
    s := s+correctlength(rsql.findfield('T_REGIMETVA').AsString, 3);
    s := s+correctlength(rsql.findfield('T_SOUMISTPF').AsString, 1);
    s := s+correctlength(rsql.findfield('T_LETTRABLE').AsString, 1);
    s := s+correctlength(rsql.findfield('T_CORRESP1').AsString, 17);
    s := s+correctlength(rsql.findfield('T_CORRESP2').AsString, 17);
    s := s+aligndroite(StrfMontant(rsql.findField('T_TOTALDEBIT').AsFloat,20,2,'',false), 20);
    s := s+alignDroite(StrfMontant(rsql.findField('T_TOTALCREDIT').AsFloat,20,2,'',false),20);
    s := s+correctLength(DateToStr(rsql.findField('T_DATECREATION').AsDateTime), 8);
    s := s+correctLength(DateToStr(rsql.findField('T_DATEMODIF').AsDateTime),8);
    s := s+correctLength(DateToStr(rsql.findField('T_DATEOUVERTURE').AsDateTime),8);
    s := s+correctLength(DateToStr(rsql.findField('T_DATEFERMETURE').AsDateTime),8);
    s := s+correctlength(rsql.findfield('T_FERME').asString, 1);
    s := s+correctLength(DateToStr(rsql.findField('T_DATEDERNMVT').AsDateTime),8);
    s := s+aligndroite(StrfMontant(rsql.findField('T_DEBITDERNMVT').AsFloat,20,2,'',false),20);
    s := s+aligndroite(StrfMontant(rsql.findField('T_CREDITDERNMVT').AsFloat,20,2,'',false),20);
    s := s+aligndroite(StrfMontant(rsql.findField('T_NUMDERNMVT').AsFloat,10,0,'',false),10);
    s := s+aligndroite(StrfMontant(rsql.findField('T_LIGNEDERNMVT').AsFloat,10,0,'',false),10);
    s := s+correctlength(rsql.findfield('T_CONFIDENTIEL').AsString, 1);
    s := s+correctlength(rsql.findfield('T_SOLDEPROGRESSIF').AsString, 1);
    s := s+correctlength(rsql.findfield('T_SAUTPAGE').AsString, 1);
    s := s+correctlength(rsql.findfield('T_TOTAUXMENSUELS').AsString, 1);
    s := s+alignDroite(StrfMontant(Rsql.FindField('T_COUTHORAIRE').AsFloat,20,2,'',false),20);
    s := s+correctlength(rsql.findfield('T_SOCIETE').AsString, 3);
    s := s+correctlength(rsql.findfield('T_RELANCEREGLEMENT').AsString, 3);
    s := s+correctlength(rsql.findfield('T_RELANCETRAITE').AsString, 3);
    s := s+correctlength(rsql.findfield('T_RESIDENTETRANGER').AsString, 3);
    s := s+correctlength(rsql.findfield('T_NATUREECONOMIQUE').AsString, 3);
    s := s+correctlength(rsql.findfield('T_MOTIFVIREMENT').AsString, 3);
    s := s+correctlength(rsql.findfield('T_LETTREPAIEMENT').AsString, 3);
    s := s+correctlength(rsql.findfield('T_RELEVEFACTURE').AsString, 1);
    s := s+correctlength(rsql.findfield('T_FREQRELEVE').AsString, 3);
    s := s+alignDroite(StrfMontant(rsql.findField('T_JOURRELEVE').AsFloat,10,0,'',false),10);
    s := s+correctlength(rsql.findfield('T_UTILISATEUR').AsString, 3);
    s := s+correctlength(rsql.findfield('T_RVA').AsString, 250);
    s := s+correctlength(rsql.findfield('T_EMAIL').AsString, 250);
    s := s+correctlength(rsql.findfield('T_DERNLETTRAGE').AsString, 4);
    s := s+correctlength(rsql.findfield('T_JURIDIQUE').AsString, 3);
    s := s+correctlength(rsql.findfield('T_FORMEJURIDIQUE').AsString, 3);
    s := s+correctlength(rsql.findfield('T_PROFIL').AsString, 3);
    s := s+correctlength(rsql.findfield('T_TVAENCAISSEMENT').AsString, 3);
    s := s+alignDroite(StrfMontant(rsql.findField('T_TOTDEBP').AsFloat,20,2,'',false),20);
    s := s+alignDroite(StrfMontant(rsql.findField('T_TOTCREP').AsFloat,20,2,'',false),20);
    s := s+alignDroite(StrfMontant(rsql.findField('T_TOTDEBE').AsFloat,20,2,'',false),20);
    s := s+alignDroite(StrfMontant(rsql.findField('T_TOTCREE').AsFloat,20,2,'',false),20);
    s := s+alignDroite(StrfMontant(rsql.findField('T_TOTDEBS').AsFloat,20,2,'',false),20);
    s := s+alignDroite(StrfMontant(rsql.findField('T_TOTCRES').AsFloat,20,2,'',false),20);
    s := s+alignDroite(StrfMontant(rsql.findField('T_TOTDEBANO').AsFloat,20,2,'',false),20);
    s := s+alignDroite(StrfMontant(rsql.findField('T_TOTCREANO').AsFloat,20,2,'',false),20);
    s := s+alignDroite(StrfMontant(rsql.findField('T_TOTDEBANON1').AsFloat,20,2,'',false),20);
    s := s+alignDroite(StrfMontant(rsql.findField('T_TOTCREANON1').AsFloat,20,2,'',false),20);
    s := s+correctLength(DateToStr(rsql.findField('T_DATEDERNRELEVE').AsDateTime),8);
    s := s+alignDroite(StrfMontant(rsql.findField('T_SCORERELANCE').AsFloat,10,0,'',false),10);
    s := s+correctlength(rsql.findfield('T_CREERPAR').AsString, 3);
    s := s+correctlength(rsql.findfield('T_EXPORTE').AsString, 3);
    s := s+alignDroite(StrfMontant(rsql.findField('T_SCORECLIENT').AsFloat,10,0,'',false),10);
    s := s+correctlength(rsql.findfield('T_PAYEURECLATEMENT').AsString, 1);
    s := s+correctLength(DateToStr(rsql.findField('T_DATEDERNPIECE').AsDateTime), 8);
    s := s+alignDroite(StrfMontant(rsql.findField('T_NUMDERNPIECE').AsFloat,10,0,'',false),10);
    s := s+alignDroite(StrfMontant(rsql.findField('T_TOTDERNPIECE').AsFloat,20,2,'',false),20);
    s := s+correctlength(rsql.findfield('T_TABLE0').AsString, 17);
    s := s+correctlength(rsql.findfield('T_TABLE1').AsString, 17);
    s := s+correctlength(rsql.findfield('T_TABLE2').AsString, 17);
    s := s+correctlength(rsql.findfield('T_TABLE3').AsString, 17);
    s := s+correctlength(rsql.findfield('T_TABLE4').AsString, 17);
    s := s+correctlength(rsql.findfield('T_TABLE5').AsString, 17);
    s := s+correctlength(rsql.findfield('T_TABLE6').AsString, 17);
    s := s+correctlength(rsql.findfield('T_TABLE7').AsString, 17);
    s := s+correctlength(rsql.findfield('T_TABLE8').AsString, 17);
    s := s+correctlength(rsql.findfield('T_TABLE9').AsString, 17);
    s := s+correctlength(rsql.findfield('T_CONSO').AsString, 3);
    s := s+alignDroite(StrfMontant(rsql.findField('T_CREDITDEMANDE').AsFloat,20,2,'',false),20);
    s := s+alignDroite(StrfMontant(rsql.findField('T_CREDITACCORDE').AsFloat,20,2,'',false),20);
    s := s+correctlength(rsql.findfield('T_DOSSIERCREDIT').AsString, 35);
    s := s+correctLength(DateToStr(rsql.findField('T_DATECREDITDEB').AsDateTime), 8);
    s := s+correctLength(DateToStr(rsql.findField('T_DATECREDITFIN').AsDateTime), 8);
    s := s+alignDroite(StrfMontant(rsql.findField('T_CREDITPLAFOND').AsFloat,20,2,'',false),20);
    s := s+correctLength(DateToStr(rsql.findField('T_DATEPLAFONDDEB').AsDateTime), 8);
    s := s+correctLength(DateToStr(rsql.findField('T_DATEPLAFONDFIN').AsDateTime), 8);
    s := s+correctlength(rsql.findfield('T_NIVEAURISQUE').AsString, 3);
    s := s+correctlength(rsql.findfield('T_AVOIRRBT').AsString, 1);
    s := s+correctlength(rsql.findfield('T_ISPAYEUR').AsString, 1);
    s := s+correctlength(rsql.findfield('T_CODEIMPORT').AsString, 35);
    s := s+correctlength(rsql.findfield('T_TIERS').AsString, 17);
    s := s+correctlength(rsql.findfield('T_NATIONALITE').AsString, 3);
    s := s+correctlength(rsql.findfield('T_PRENOM').AsString, 35);
    s := s+alignDroite(StrfMontant(rsql.findField('T_JOURNAISSANCE').AsFloat,10,0,'',false),10);
    s := s+alignDroite(StrfMontant(rsql.findField('T_MOISNAISSANCE').AsFloat,10,0,'',false),10);
    s := s+alignDroite(StrfMontant(rsql.findField('T_ANNEENAISSANCE').AsFloat,10,0,'',false),10);
    s := s+correctlength(rsql.findfield('T_COMPTATIERS').AsString, 3);
    s := s+correctlength(rsql.findfield('T_DEBRAYEPAYEUR').AsString, 1);
    s := s+alignDroite(StrfMontant(rsql.findField('T_MOISCLOTURE').AsFloat,10,0,'',false),10);
    s := s+correctlength(rsql.findfield('T_EURODEFAUT').AsString, 1);
    s := s+correctlength(rsql.findfield('T_PARTICULIER').AsString, 1);
    s := s+correctlength(rsql.findfield('T_SEXE').AsString, 3);
    s := s+correctlength(rsql.findfield('T_PASSWINTERNET').AsString, 20);
    s := s+correctlength(rsql.findfield('T_PRESCRIPTEUR').AsString, 17);
    s := s+correctlength(rsql.findfield('T_PUBLIPOSTAGE').AsString, 1);
    s := s+correctlength(rsql.findfield('T_ORIGINETIERS').AsString, 3);
    s := s+correctlength(rsql.findfield('T_SOCIETEGROUPE').AsString, 17);
    s := s+correctlength(rsql.findfield('T_PHONETIQUE').AsString, 35);
    s := s+correctLength(DateToStr(rsql.findField('T_DATEPROCLI').AsDateTime), 8);
    s := s+correctlength(rsql.findfield('T_ETATRISQUE').AsString, 3);
    s := s+correctlength(rsql.findfield('T_DOMAINE').AsString, 3);
    s := s+correctLength(DateToStr(rsql.findField('T_DATEINTEGR').AsDateTime), 8);
    s := s+correctlength(rsql.findfield('T_NIVEAUIMPORTANCE').AsString, 3);
    s := s+correctlength(rsql.findfield('T_CLETELEPHONE').AsString, 25);
    s := s+correctlength(rsql.findfield('T_ENSEIGNE').AsString, 70);
    s := s+correctlength(rsql.findfield('T_EMAILING').AsString, 1);
    s := s+correctlength(rsql.findfield('T_REGION').AsString, 9);
    s := s+alignDroite(StrfMontant(rsql.findField('T_DELAIMOYEN').AsFloat,10,0,'',false),10);
    end;
    writeln(f, s);
    ShortDateFormat := formatDate;
end;
{------------------------------------------------------------------------------}
{Génère un fichier TRA contenant la table voulue}
function TImportDP.generateFile(nomTable : string) : integer;
var
  Rsql : TQuery;
  f : TextFile;
begin
  assignFile(f, 'C:/generation.tra');
  rewrite(f);
  try
    Rsql := OpenSql('SELECT * FROM '+nomTable, true);
  except
    showmessage('Héhé, erreur pauv'' tache !');
  end;
  While Not Rsql.Eof do//parcours des résultats
    begin
    writeLigne(rsql, f, nomTable);
    Rsql.Next;
    end;
  closeFile(f);
  result := 0;
  rsql.free;
end;

end.
