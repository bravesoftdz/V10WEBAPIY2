unit uLibExportTVA;

interface

uses classes,
     ParamSoc,
     Sysutils,
     HCtrls,
     HEnt1,
     HmsgBox,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     uTOB,
     uTOBDebug;
var
    MESS : array [1..7] of string = (
        {1}'Le Filtre ne correspond à aucun enregistrement.',
        {2}'Le montant ne peut etre superieur à ''9 999 999 999''.',
        {3}'Le montant ne peut etre inferieur à ''- 9 999 999 999''.',
        {4}'Veuillez renseigner le paramètre société : SO_DECLATVA.',
        {5}'Numéro de TVA vide.',
        {6}'Trimestre vide.',
        {7}'Année vide.'
        );
procedure LanceTraitement(NumDeclaration : integer ; Modifiable : boolean);
procedure DeleteModifsDeclaration(NumDeclaration : integer);
procedure DeleteModif(NumModif : integer);
procedure InsertModif(Modification : integer; Declaration : integer);

procedure PutLigne(var Ligne : string; pos : integer; rajout : string);
function ArrondirApresVirgule(Nombre : double; NbChiffre : integer) : Double;

function GetDeclarationLibelle(     NumDeclaration : Integer)   : String;
function IsModifEditable(           NumDeclaration : Integer)   : boolean;
function DeclarationHaveModif(      NumDeclaration : Integer)   : boolean;
function GetTypeDeclaration(        NumDeclaration : Integer)   : String;
function GetDateValidDeclaration(   NumDeclaration : Integer)   : TDateTime; 
function GetDebutExercice(          NumDeclaration : Integer)   : TDateTime;
function GetFinExercice(            NumDeclaration : Integer)   : TDateTime;
function GetNatureDeclaration(      NumDeclaration : Integer)   : string;
function GetTypeColDeclaration(     NumDeclaration : Integer)   : string;
function GetRegimeTVADeclaration(   NumDeclaration : Integer)   : string;
function GetAnneeDeclaration(       NumDeclaration : Integer)   : string;
function GetTrimDeclaration(        NumDeclaration : Integer)   : string;

function GetCodeMax(Table : String) : integer;
function ReplaceChar(Combos : String) : string;
function DeleteChar(Source : String;Car : Char) : string;
function CompleteWithCar(Chaine : string ; Car : Char ; Long : integer; sens : Char) : string;
function GetMinAnneeExercice : integer ;
function GetMaxAnneeExercice : integer ;

function LancerSaveDialog(var Repertoire : string) : Boolean;

type
    TDeclaration = class
    public
      Numero          : integer;
      Libelle         : string;
      Annee           : string;
      Trimestre       : string;
      TypeDeclaration : string;
      DateValid       : TDateTime;
      DebutPeriode    : TDateTime;
      FinPeriode      : TDateTime;
      Nature          : string;
      CompteCollectif : string;
      RegimeTVA       : string;
      isValide        : boolean;
      haveModif       : boolean;
      isOK            : boolean;
	   ErrorMsg		    : integer;
      Enregistrements : TOB;
      Modifications	 : TOB; 
      Recapitulatif	 : TOB;
      Traitement      : TOB;
      MontantTVA		 : double;
      MontantHT		 : double;
      NbEnreg			 : integer;
      constructor Create(NumDeclaration : Integer);
      destructor  Destroy;override;
      procedure   Valide;
      procedure   MakeTOB;
    private
      Lignes :  TStringList;
      procedure TraiteEntete();
      procedure TraiteDetail();
      procedure TraiteModifs();
      procedure TraiteFin();
      procedure ChargeTOB();
      procedure ChargeTOBModif();
      procedure ChargeRecap();
    end;

implementation

uses utilPGI
     ,dialogs
{$IFNDEF EAGLCLIENT}
     , EdtREtat
{$ELSE}
     , UtileAgl
{$ENDIF}
;

constructor TDeclaration.Create(NumDeclaration : Integer);
begin
   Numero          := NumDeclaration;
   Libelle         := GetDeclarationLibelle(NumDeclaration);
   Annee           := GetAnneeDeclaration(NumDeclaration);
   Trimestre       := GetTrimDeclaration(NumDeclaration);
   TypeDeclaration := GetTypeDeclaration(NumDeclaration);
   DateValid       := GetDateValidDeclaration(NumDeclaration);
   DebutPeriode    := GetDebutExercice(NumDeclaration);
   FinPeriode      := GetFinExercice(NumDeclaration);
   Nature          := GetNatureDeclaration(NumDeclaration);
   { BVE 15/03/2007 FQ 19777
   TypeCollectif   := GetTypeColDeclaration(NumDeclaration);
   CompteCollectif := GetCompteColDeclaration(NumDeclaration);
   } 
   CompteCollectif := GetTypeColDeclaration(NumDeclaration);
   RegimeTVA       := GetRegimeTVADeclaration(NumDeclaration);
   isValide        := not(IsModifEditable(NumDeclaration));
   haveModif       := DeclarationHaveModif(NumDeclaration);
   isOK            := true;
   ErrorMsg		    := 0;
   Lignes          := TStringList.Create;
   Enregistrements := TOB.Create('Ma Declaration', nil, -1);
   Modifications   := TOB.Create('HISTOMODIFTVA', nil, -1);
   Recapitulatif   := TOB.Create('Recapitulatif',nil, -1);
   Traitement      := TOB.Create('Mon Traitement',nil, -1);
   MontantTVA		 := 0;
   MontantHT		 := 0;
   NbEnreg			 := 0;
   // Chargement de la TOB
   ChargeTOB;
   // Chargement des modifs
	ChargeTOBModif;
   // Chargement de la TOB Traitement;
   MakeTOB; 
end;

destructor TDeclaration.Destroy;
begin
  inherited;
  if assigned(Lignes)           then FreeAndNil(Lignes);
  if assigned(Enregistrements)  then FreeAndNil(Enregistrements);
  if assigned(Modifications)    then FreeAndNil(Modifications);
  if assigned(Recapitulatif)    then FreeAndNil(Recapitulatif);
  if assigned(Traitement)       then FreeAndNil(Traitement);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/01/2007
Modifié le ... :   /  /    
Description .. : Procedure de traitement de la déclaration
Mots clefs ... : 
*****************************************************************}
procedure LanceTraitement(NumDeclaration : integer ; Modifiable : boolean);
var
   maDeclaration : TDeclaration;
   FileName      : string;
begin

   maDeclaration := TDeclaration.Create(NumDeclaration);
   try
      {FileName := GetParamSocSecur('SO_DECLATVA', '');
      if FileName = '' then
      begin
         maDeclaration.isOK := false;
         maDeclaration.ErrorMsg := 4;
      end;    }
      maDeclaration.isOK := LancerSaveDialog(Filename);
      if maDeclaration.isOK then maDeclaration.TraiteEntete;
      if maDeclaration.isOK then maDeclaration.TraiteDetail;
      if maDeclaration.isOK then maDeclaration.TraiteModifs;
      if maDeclaration.isOK then maDeclaration.TraiteFin;

      // Si tout c'est bien passée on valide la déclaration.
      if maDeclaration.isOK then
      begin
         // On sauvegarde le fichier
         maDeclaration.Lignes.SaveToFile(FileName);
         maDeclaration.ChargeRecap;
         if maDeclaration.TypeDeclaration = '723' then
            LanceEtatTOB('E','BEL','NO3',maDeclaration.Recapitulatif ,true,false,false,nil,'',TraduireMemoire('Note récapitulative'),false)
         else if maDeclaration.TypeDeclaration = '725' then
            LanceEtatTOB('E','BEL','NO5',maDeclaration.Recapitulatif ,true,false,false,nil,'',TraduireMemoire('Note récapitulative'),false);
         if Modifiable then
         begin
         { FQ 19775 BVE 08/03/2007
            // On s'occuppe du récapitulatif :
            maDeclaration.ChargeRecap;
            LanceEtatTOB('E','BEL','NOT',maDeclaration.Recapitulatif ,true,false,false,nil,'',TraduireMemoire('Note récapitulative'),false);
          }
            PGIInfo('La déclaration a été générée dans ' + #10#13 + FileName,maDeclaration.Libelle);
            maDeclaration.Valide;
         end
         else
            PGIInfo('La déclaration a été rééditée dans ' + #10#13 + FileName,maDeclaration.Libelle);
      end
      else
      begin
      	// Il y a eu une erreur
         if maDeclaration.ErrorMsg <> 0 then
            PGIError(MESS[maDeclaration.ErrorMsg],maDeclaration.Libelle);
      end;
   finally
      FreeAndNil(maDeclaration);
   end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : Procedure permettant de charger la TOB avec l'ensemble 
Suite ........ : des valeurs necessaire à l'édition
Mots clefs ... :
*****************************************************************}
procedure TDeclaration.ChargeTOB();
var
  SQL : string;
  Q   : TQuery;
  i   : integer;
begin
   {FQ 19783 BVE 13/03/2007}
   {SQL := 'SELECT E_CONTREPARTIEAUX, E_TYPEMVT, SUM(E_CREDIT - E_DEBIT) as MONTANT, ' +
          'MAX(T_PAYS) AS PAYS, MAX(T_NIF) AS NUMTVA ' +
          'FROM ECRITURE, TIERS ' +
          'WHERE E_CONTREPARTIEAUX = T_AUXILIAIRE AND LEN(TRIM(T_NIF)) > 0 ' +
          'AND E_DATECOMPTABLE >= "' + UsDateTime(DebutPeriode) + '"  AND E_DATECOMPTABLE <= "' + UsDateTime(FinPeriode) + '" ' +
          'AND (E_TYPEMVT="HT" or E_TYPEMVT="TVA" )';                                                                            }
   SQL := 'SELECT E_CONTREPARTIEAUX, E_TYPEMVT, SUM(E_CREDIT - E_DEBIT) as MONTANT, ' +
          'MAX(T_NIF) AS NUMTVA ' +
          'FROM ECRITURE, TIERS ' +
          'WHERE E_CONTREPARTIEAUX = T_AUXILIAIRE AND LEN(TRIM(T_NIF)) > 0 ' +
          'AND E_DATECOMPTABLE >= "' + UsDateTime(DebutPeriode) + '"  AND E_DATECOMPTABLE <= "' + UsDateTime(FinPeriode) + '" ' +
          'AND (E_TYPEMVT="HT" or E_TYPEMVT="TVA" )'; 
   {END FQ 19783 BVE 13/03/2007}
	if ReplaceChar(Nature) <> '""' then
		SQL := SQL + ' AND E_NATUREPIECE IN (' + ReplaceChar(Nature) + ')';
   {BVE 15/03/2007 Suppression Compte Collectif suite FQ19777 demande ED}
  	if ReplaceChar(CompteCollectif) <> '""' then
		SQL := SQL + ' AND T_COLLECTIF IN (' + ReplaceChar(CompteCollectif) + ')';
   {
   if ReplaceChar(CompteCollectif) <> '""' then
		SQL := SQL + ' AND T_NATUREAUXI IN (' + ReplaceChar(CompteCollectif) + ')';
   END BVE 15/03/2007}
	if ReplaceChar(RegimeTVA) <> '""' then
		SQL := SQL + ' AND T_REGIMETVA IN (' + ReplaceChar(RegimeTVA) + ')';

   SQL := SQL + ' GROUP BY E_CONTREPARTIEAUX, E_TYPEMVT ORDER BY E_CONTREPARTIEAUX';
	Q := OpenSQL(SQL,true);
	if not(Q.Eof) then
		enregistrements.LoadDetailDB('Ma Déclaration','','',Q,false)
	else
	begin
   	isOK := false;
      ErrorMsg := 1;
  	end;

   // On remplace le code ISO 3 par le code ISO 2.
   // On verifie que le numéro TVA est bien un numero
   // et ne comporte pas l'ISO du pays :

   for i := 0 to enregistrements.detail.count - 1 do
   begin
      {if Trim(Enregistrements.Detail[i].GetValue('NUMTVA')) <> '' then
      begin
         if (UpperCase(Enregistrements.Detail[i].GetValue('NUMTVA'))[1] in ['A'..'Z']) and
            (UpperCase(Enregistrements.Detail[i].GetValue('NUMTVA'))[2] in ['A'..'Z']) then
         begin
            Enregistrements.Detail[i].PutValue('PAYS',UpperCase(Copy(Enregistrements.Detail[i].GetValue('NUMTVA'),1,2)));
            Enregistrements.Detail[i].PutValue('NUMTVA',Copy(Enregistrements.Detail[i].GetValue('NUMTVA'),3,10));
         end
         else if (UpperCase(Enregistrements.Detail[i].GetValue('NUMTVA'))[1] = 'B') then
         begin
            Enregistrements.Detail[i].PutValue('PAYS','BE');
            Enregistrements.Detail[i].PutValue('NUMTVA',Copy(Enregistrements.Detail[i].GetValue('NUMTVA'),2,10));
         end;
      end;
      if (Enregistrements.Detail[i].GetValue('PAYS') <> '') and (length(Enregistrements.Detail[i].GetValue('PAYS')) > 2) then
         Enregistrements.Detail[i].PutValue('PAYS',codeISODuPays(Enregistrements.Detail[i].GetValue('PAYS')));}
      Enregistrements.Detail[i].AddChampSup('PAYS',false);
      if Length(Trim(Enregistrements.Detail[i].GetValue('NUMTVA'))) > 2 then
      begin
         Enregistrements.Detail[i].PutValue('PAYS',UpperCase(Copy(Enregistrements.Detail[i].GetValue('NUMTVA'),1,2)));
         Enregistrements.Detail[i].PutValue('NUMTVA',Copy(Enregistrements.Detail[i].GetValue('NUMTVA'),3,Length(Trim(Enregistrements.Detail[i].GetValue('NUMTVA'))) - 2 ));
      end
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 31/01/2007
Modifié le ... :   /  /    
Description .. : Procedure permettant de remplir la TOB qui servira à editer 
Suite ........ : le récapitulatif
Mots clefs ... : 
*****************************************************************}
procedure TDeclaration.ChargeRecap;
Var
  Fille : TOB;
begin
  Fille := TOB.Create('Recapitulatif',Recapitulatif,0);
  Fille.AddChampSup('TRIMESTRE',true);
  Fille.AddChampSup('ANNEE',true);
  Fille.AddChampSup('RECORD',true);
  Fille.AddChampSup('MONTANT_CA',true);
  Fille.AddChampSup('MONTANT_TVA',true);
  Fille.AddChampSup('DATE',true);
  Fille.PutValue('TRIMESTRE',Trimestre);
  Fille.PutValue('ANNEE',Annee);
  Fille.PutValue('RECORD',NbEnreg);
  Fille.PutValue('MONTANT_CA',MontantHT);
  Fille.PutValue('MONTANT_TVA',MontantTVA);
  Fille.PutValue('DATE',now);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : Procedure permettant de charger la TOB avec l'ensemble 
Suite ........ : des valeurs necessaire à l'édition
Mots clefs ... :
*****************************************************************}
procedure TDeclaration.ChargeTOBModif();
var
  SQL : string;
  Q   : TQuery;
begin
   SQL := 'SELECT * ' +
          'FROM HISTOMODIFTVA ' +
          'WHERE HMT_CODECPT = "' + IntToStr(Numero) + '"';
	Q := OpenSQL(SQL,true);
	if not(Q.Eof) then
		Modifications.LoadDetailDB('HISTOMODIFTVA','','',Q,false)
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/01/2007
Modifié le ... :   /  /    
Description .. : Procedure permettant de gerer les lignes d'entetes
Suite ........ : (type 555555 et 000000)
Mots clefs ... : 
*****************************************************************}
procedure TDeclaration.TraiteEntete();
var
   entete : string;
begin
   // Traitement de la premiere ligne d'entete. (555555)
   PutLigne(Entete,1,'555555');
   PutLigne(Entete,7,GetParamSocSecur('SO_LIBELLE', ''));
   PutLigne(Entete,39,GetParamSocSecur('SO_ADRESSE1', ''));
   PutLigne(Entete,63,GetParamSocSecur('SO_CODEPOSTAL', '') + ' ' + GetParamSocSecur('SO_VILLE', ''));
   if TypeDeclaration = '723' then
      PutLigne(Entete,119,CompleteWithCar('1','0',10,'G'))
   else if TypeDeclaration = '725' then
      PutLigne(Entete,101,CompleteWithCar('1','0',10,'G'));
   // On force la longueur de l'enregistrement.
   Lignes.Add(CompleteWithCar(Entete,' ',128,'D'));

   Entete := '';
   // Traitement de la seconde ligne d'entete. (000000)
   PutLigne(Entete,1,'000000');
   PutLigne(Entete,7,GetParamSocSecur('SO_LIBELLE', ''));
   PutLigne(Entete,39,GetParamSocSecur('SO_ADRESSE1', ''));
   PutLigne(Entete,63,GetParamSocSecur('SO_CODEPOSTAL', '') + ' ' + GetParamSocSecur('SO_VILLE', ''));
   if TypeDeclaration = '723' then
   begin
      PutLigne(Entete,91,'BE');
      PutLigne(Entete,93,GetParamSocSecur('SO_RVA', ''));
      PutLigne(Entete,119,Trimestre);
      PutLigne(Entete,120,Annee); 
      PutLigne(Entete,124,'E');
   end
   else if TypeDeclaration = '725' then
   begin
      if Length(GetParamSocSecur('SO_RVA', '')) > 9 then
      begin
         PutLigne(Entete,90,'B');
         PutLigne(Entete,91,GetParamSocSecur('SO_RVA', ''));
      end
      else
      begin
         PutLigne(Entete,90,'BE');
         PutLigne(Entete,92,GetParamSocSecur('SO_RVA', ''));
      end;
      PutLigne(Entete,121,'E');
      PutLigne(Entete,122,Annee);
   end;
   Lignes.Add(CompleteWithCar(Entete,' ',128,'D'));
end;
    
{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/01/2007
Modifié le ... :   /  /    
Description .. : Procedure permettant de traiter les données issues de la
Suite ........ : TOB et de les renseigner dans le fichier
Mots clefs ... : 
*****************************************************************}
procedure TDeclaration.TraiteDetail();
var
   sequence : integer;
   i        : integer;
   Detail	: string;
begin
   i := 0;
   sequence := 1 ;
   while i < Enregistrements.Detail.count do
   begin
   	if TypeDeclaration = '723' then
      begin
      { FQ 19788 BVE 13/03/2007
      	if Enregistrements.Detail[i].GetValue('E_TYPEMVT') = 'TVA' then
        FQ 19788 BVE 13/03/2007 }
        	if Enregistrements.Detail[i].GetValue('E_TYPEMVT') = 'HT' then
         begin
         	// On ne recupere que les montants TVA !
            PutLigne(Detail,1,CompleteWithCar(IntToStr(sequence),'0',6,'G'));
            PutLigne(Detail,91,Enregistrements.Detail[i].GetValue('PAYS'));
            if Trim(Enregistrements.Detail[i].GetValue('NUMTVA')) <> '' then
            	PutLigne(Detail,93,Enregistrements.Detail[i].GetValue('NUMTVA'))
            else
               PutLigne(Detail,93,'000000000');
            PutLigne(Detail,106,CompleteWithCar(GetValueMontant(Enregistrements.Detail[i].GetValue('MONTANT')),'0',13,'G'));
            Lignes.Add(CompleteWithCar(Detail,' ',128,'D'));
            Detail := '';
            Inc(sequence);

            // On sauvegarde le montant
            MontantHT := MontantHT + StrToFloat(Enregistrements.Detail[i].GetValue('MONTANT'));
         end;
      end
      else if TypeDeclaration = '725' then
      begin
         PutLigne(Detail,1,CompleteWithCar(IntToStr(sequence),'0',6,'G'));
         if Length(Enregistrements.Detail[i].GetValue('NUMTVA')) = 10 then
         begin
         	PutLigne(Detail,90,'B');
            PutLigne(Detail,91,Enregistrements.Detail[i].GetValue('NUMTVA'));
         end
         else
         begin
            PutLigne(Detail,90,'BE');
            if Trim(Enregistrements.Detail[i].GetValue('NUMTVA')) <> '' then
            	PutLigne(Detail,92,Enregistrements.Detail[i].GetValue('NUMTVA'))
            else
               PutLigne(Detail,92,'000000000');
         end;
         if Enregistrements.Detail[i].GetValue('E_TYPEMVT') = 'TVA' then
         begin
         	PutLigne(Detail,111,CompleteWithCar(GetValueMontant(Enregistrements.Detail[i].GetValue('MONTANT')),'0',10,'G'));
				MontantTVA := MontantTVA + StrToFloat(Enregistrements.Detail[i].GetValue('MONTANT'));
            // On verifie l'enregistrement suivant pour recuperer le HT
            if (i + 1) < Enregistrements.Detail.count then
            begin
            	if Enregistrements.Detail[i].GetValue('E_CONTREPARTIEAUX') = Enregistrements.Detail[i+1].GetValue('E_CONTREPARTIEAUX') then
	            begin
   	         	if Enregistrements.Detail[i+1].GetValue('E_TYPEMVT') = 'HT' then
      	         begin
	      	         PutLigne(Detail,101,CompleteWithCar(GetValueMontant(Enregistrements.Detail[i+1].GetValue('MONTANT')),'0',10,'G'));
							MontantHT := MontantHT + StrToFloat(Enregistrements.Detail[i+1].GetValue('MONTANT'));
               	   Inc(i);
		            end;
      	      end;
            end;
         end
         else if Enregistrements.Detail[i].GetValue('E_TYPEMVT') = 'HT' then
         begin
         	PutLigne(Detail,101,CompleteWithCar(GetValueMontant(Enregistrements.Detail[i].GetValue('MONTANT')),'0',10,'G'));
				MontantHT := MontantHT + StrToFloat(Enregistrements.Detail[i].GetValue('MONTANT'));
            // On verifie l'enregistrement suivant pour recuperer la TVA
            if (i+1) < Enregistrements.Detail.count then
            begin
	            if Enregistrements.Detail[i].GetValue('E_CONTREPARTIEAUX') = Enregistrements.Detail[i+1].GetValue('E_CONTREPARTIEAUX') then
   	         begin
      	      	if Enregistrements.Detail[i+1].GetValue('E_TYPEMVT') = 'TVA' then
         	      begin
	         	      PutLigne(Detail,111,CompleteWithCar(GetValueMontant(Enregistrements.Detail[i+1].GetValue('MONTANT')),'0',10,'G'));
							MontantTVA := MontantTVA + StrToFloat(Enregistrements.Detail[i+1].GetValue('MONTANT'));
                  	Inc(i);
	               end;
   	         end;
      	   end;
         	Lignes.Add(CompleteWithCar(Detail,' ',128,'D'));
	         Detail := '';
   	      Inc(sequence);
      	end;
      end;
      Inc(i);
   end;
   NbEnreg := sequence - 1;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/01/2007
Modifié le ... :   /  /    
Description .. : Procedure permettant de rajouter les modifications à la 
Suite ........ : déclaration.
Mots clefs ... : 
*****************************************************************}
procedure TDeclaration.TraiteModifs();
var
	i : integer;
   modif : string;
begin
   i := 0;
   { FQ 19788 BVE 16/03/2007 }
   while i < Modifications.Detail.count do
   begin
      Inc(NbEnreg);
      PutLigne(Modif,1,CompleteWithCar(IntToStr(NbEnreg),'0',6,'G'));
      PutLigne(Modif,91,codeISODuPAys(Modifications.Detail[i].GetValue('HMT_CODEPAYS')));
      PutLigne(Modif,93,Modifications.Detail[i].GetValue('HMT_NUMTVA'));
      PutLigne(Modif,105,Modifications.Detail[i].GetValue('HMT_CODEMODIF'));
      PutLigne(Modif,106,CompleteWithCar(GetValueMontant(Modifications.Detail[i].GetValue('HMT_MONTANT')),'0',13,'G'));
      //SDA le 02/01/2008
      MontantHT := MontantHT + StrToFloat(Modifications.Detail[i].GetValue('HMT_MONTANT'));
      //Fin SDA le 02/01/2008
      PutLigne(Modif,119,Modifications.Detail[i].GetValue('HMT_TRIMESTRE'));
      PutLigne(Modif,120,Modifications.Detail[i].GetValue('HMT_ANNEE'));
      Lignes.Add(CompleteWithCar(Modif,' ',128,'D'));
      Modif := '';
   	Inc(i);
   end;
   { END FQ 19788 BVE 16/03/2007 }
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/01/2007
Modifié le ... :   /  /    
Description .. : Procedure gérant la ligne de fin récapitulative.
Mots clefs ... : 
*****************************************************************}
procedure TDeclaration.TraiteFin();
var
	Fin	: string;
begin
	if TypeDeclaration = '723' then
   begin
   	PutLigne(Fin,1,'999999');
      PutLigne(Fin,91,'BE');
      PutLigne(Fin,93,GetParamSocSecur('SO_RVA', ''));
      PutLigne(Fin,106,CompleteWithCar(GetValueMontant(MontantHT),'0',13,'G'));
      PutLigne(Fin,119,CompleteWithCar(IntToStr(NbEnreg),'0',6,'G'));
   end
   else if TypeDeclaration = '725' then
	begin
   	PutLigne(Fin,1,'999999');                          
      PutLigne(Fin,7,CompleteWithCar(GetValueMontant(MontantHT),'0',16,'G'));
      PutLigne(Fin,23,CompleteWithCar(GetValueMontant(MontantTVA),'0',16,'G'));
      if length(GetParamSocSecur('SO_RVA', '')) = 10 then
      begin
	      PutLigne(Fin,90,'B');
   	   PutLigne(Fin,91,GetParamSocSecur('SO_RVA', ''));
      end
      else
      begin
	      PutLigne(Fin,90,'BE');
   	   PutLigne(Fin,92,GetParamSocSecur('SO_RVA', ''));
      end;
   end;
   Lignes.Add(CompleteWithCar(Fin,' ',128,'D'));

end;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/01/2007
Modifié le ... :   /  /    
Description .. : Procedure permettant de valider une déclaration
Mots clefs ... : 
*****************************************************************}
procedure TDeclaration.Valide();
begin
   ExecuteSQL('UPDATE PARAMEXTTVA SET CPT_DATEVALID="' + USDateTime(now) + '" WHERE CPT_CODE="' + IntToStr(Numero) + '"');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 14/03/2007
Modifié le ... :   /  /    
Description .. : Procedure permettant de créer la TOB Traitement qui 
Suite ........ : contient toutes les informations pour toutes les déclarations
Mots clefs ... : 
*****************************************************************}
procedure TDeclaration.MakeTOB();
var
  i       : integer;
  Dst     : TOB;
  Src     : TOB;

procedure _MakeChamp(var Fille : TOB);
begin
// Formatage de la TOB :
  //SDA le 17/12/2007 version belge
  Fille.AddChampSup('CO_TYPE',true);
  //Fin SDA le 17/12/2007
  // Num Auxiliaire
  Fille.AddChampSup('E_CONTREPARTIEAUX',true);     
  // Pays
  Fille.AddChampSup('PAYS',true);
  // Num TVA
  Fille.AddChampSup('T_NIF',true);    
  // Code Motif
  Fille.AddChampSup('CODE_MODIF',true);
  // Montant HT
  Fille.AddChampSup('MONTANT_CA',true);
  // Montant TVA
  Fille.AddChampSup('MONTANT_TVA',true);
  // Trimestre Modif
  Fille.AddChampSup('TRIMESTRE_MODIF',true);
  // Année Modif
  Fille.AddChampSup('ANNEE_MODIF',true);
  // Trimestre Déclaration
  Fille.AddChampSup('TRIMESTRE',true);
  // Année
  Fille.AddChampSup('ANNEE',true);
end;
begin

  // Chargement enregistrements
  i := 0;
  while (i < Enregistrements.Detail.Count) do
  begin
     Dst := TOB.Create('',Traitement,-1);
     _MakeChamp(Dst);
     Src := Enregistrements.Detail[i];

     Dst.PutValue('E_CONTREPARTIEAUX',Src.GetValue('E_CONTREPARTIEAUX'));
     Dst.PutValue('T_NIF',Src.GetValue('NUMTVA'));
     if Src.GetValue('E_TYPEMVT') = 'HT' then
     begin
        Dst.PutValue('MONTANT_CA',Src.GetValue('MONTANT'));
        // On recherche le suivant
        if ( (i + 1) <  Enregistrements.Detail.Count) then
        begin
           if Src.GetValue('E_CONTREPARTIEAUX') = Enregistrements.Detail[i+1].GetValue('E_CONTREPARTIEAUX') then
           begin
              if Enregistrements.Detail[i+1].GetValue('E_TYPEMVT') = 'TVA' then
              begin
                 Dst.PutValue('MONTANT_TVA',Enregistrements.Detail[i+1].GetValue('MONTANT'));
                 Inc(i);
              end;
           end;
        end;
     end
     else if Src.GetValue('E_TYPEMVT') = 'TVA' then
     begin
        Dst.PutValue('MONTANT_TVA',Src.GetValue('MONTANT'));
        // On recherche le suivant
        if ( (i + 1) <  Enregistrements.Detail.Count) then
        begin
           if Src.GetValue('E_CONTREPARTIEAUX') = Enregistrements.Detail[i+1].GetValue('E_CONTREPARTIEAUX') then
           begin
              if Enregistrements.Detail[i+1].GetValue('E_TYPEMVT') = 'HT' then
              begin
                 Dst.PutValue('MONTANT_CA',Enregistrements.Detail[i+1].GetValue('MONTANT'));
                 Inc(i);
              end;
           end;
        end;
     end;
     Dst.PutValue('TRIMESTRE_MODIF','');
     Dst.PutValue('ANNEE_MODIF','');
     Dst.PutValue('CODE_MODIF','');
     Dst.PutValue('TRIMESTRE',Trimestre);
     Dst.PutValue('ANNEE',Annee);     
     Dst.PutValue('PAYS',Src.GetValue('PAYS'));
     //SDA le 17/12/2007 version belge
     Dst.PutValue('CO_TYPE','');
     //Fin SDA le 17/12/2007
     Inc(i);
  end;

  // Chargement Modif :
  i := 0;
  while (i < Modifications.Detail.Count) do
  begin    
     Dst := TOB.Create('',Traitement,-1); 
     _MakeChamp(Dst);
     Src := Modifications.Detail[i];
     
     Dst.PutValue('E_CONTREPARTIEAUX','');
     Dst.PutValue('T_NIF',Src.GetValue('HMT_NUMTVA'));
     Dst.PutValue('MONTANT_CA',Src.GetValue('HMT_MONTANT'));
     Dst.PutValue('MONTANT_TVA','');
     Dst.PutValue('TRIMESTRE_MODIF',Src.GetValue('HMT_TRIMESTRE'));
     Dst.PutValue('ANNEE_MODIF',Src.GetValue('HMT_ANNEE'));
     Dst.PutValue('CODE_MODIF',Src.GetValue('HMT_CODEMODIF'));
     Dst.PutValue('TRIMESTRE',Trimestre);
     Dst.PutValue('ANNEE',Annee);      
     Dst.PutValue('PAYS',Src.GetValue('HMT_CODEPAYS'));
     //SDA le 17/12/2007 version belge
     Dst.PutValue('CO_TYPE','');
     //Fin SDA le 17/12/2007
     Inc(i);
  end;
end;
{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : Fonction supprimant les modifications associées au code de
Suite ........ : déclaration passée en paramètre.
Mots clefs ... :
*****************************************************************}
procedure  DeleteModifsDeclaration(NumDeclaration : integer);
begin
   ExecuteSQL('DELETE FROM HISTOMODIFTVA WHERE HMT_CODECPT="' + IntToStr(NumDeclaration) + '"');
end;

procedure DeleteModif(NumModif : integer);
begin
  ExecuteSQL('DELETE FROM HISTOMODIFTVA WHERE HMT_CODE="' + IntToStr(NumModif) + '"');
end;


procedure InsertModif(Modification : integer; Declaration : integer);
begin
  ExecuteSQL('INSERT INTO HISTOMODIFTVA (HMT_CODE,HMT_CODECPT) VALUES ("' + IntToStr(Modification) + '","' + IntToStr(Declaration) + '")');
end;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : Fonction retournant le type du code de la déclaration
Suite ........ : passée en paramètre
Mots clefs ... :
*****************************************************************}
function GetTypeDeclaration(NumDeclaration : Integer) : String;
var
   SQL : String;
   Q : TQuery;
begin
   result := '';
   SQL := 'SELECT CPT_TYPE FROM PARAMEXTTVA WHERE CPT_CODE="' + IntToStr(NumDeclaration) + '"';
   try
     Q := OpenSQL(SQL,true);
     If Not Q.Eof then Result := Q.Fields[0].AsString ;
   finally
     ferme(Q);
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : Fonction retournant vrai lorsque la déclaration n'est pas 
Suite ........ : encore validée (Date de validation <> i1900) et faux sinon.
Mots clefs ... : 
*****************************************************************}
function IsModifEditable(NumDeclaration : Integer) : boolean;
var
   SQL : String;
   Q : TQuery;
begin
   SQL := 'SELECT CPT_DATEVALID FROM PARAMEXTTVA WHERE CPT_CODE="' + IntToStr(NumDeclaration) + '"';
   try
     Q := OpenSQL(SQL,true);
     If Not Q.Eof then
     begin
       if (Q.Fields[0].AsDateTime <> iDate1900) then
          Result := false
       else
          Result := true;
     end
     else
         Result := true; // Pas encore d'enregistrement.
   finally
     ferme(Q);
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : Fonction indiquant s'il y a des corrections associées au
Suite ........ : code de la déclaration passée en paramètre.
Mots clefs ... :
*****************************************************************}
function DeclarationHaveModif(NumDeclaration : Integer) : boolean;
var
   SQL : String;
begin
   SQL := 'SELECT HMT_CODE FROM HISTOMODIFTVA WHERE HMT_CODECPT="' + IntToStr(NumDeclaration) + '"';
   result := ExisteSQL(SQL);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : Fonction retournant le code le plus eleve de la table passé
Suite ........ : en parametre.
Mots clefs ... :
*****************************************************************}
function GetCodeMax(Table : String) : integer;
var
   SQL : String;
   Q : TQuery;
begin
   result := 0;
   SQL := '';
   if Table = 'PARAMEXTTVA' then
     SQL := 'SELECT MAX(CPT_CODE) FROM PARAMEXTTVA'
   else if Table = 'HISTOMODIFTVA' then
     SQL := 'SELECT MAX(HMT_CODE) FROM HISTOMODIFTVA';
   if SQL <> '' then
   begin
      try
         Q := OpenSQL(SQL,true);
         If Not Q.Eof then Result := Q.Fields[0].AsInteger ;
      finally
         ferme(Q);
      end;
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : Fonction retournant le libelle du code de la déclaration 
Suite ........ : passée en paramètre
Mots clefs ... :
*****************************************************************}
function GetDeclarationLibelle(NumDeclaration : Integer) : String;
var
   SQL : String;
   Q : TQuery;
begin
   result := '';
   SQL := 'SELECT CPT_LIBELLE FROM PARAMEXTTVA WHERE CPT_CODE="' + IntToStr(NumDeclaration) + '"';
   try
     Q := OpenSQL(SQL,true);
     If Not Q.Eof then Result := Q.Fields[0].AsString ;
   finally
     ferme(Q);
   end;
end;




{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/01/2007
Modifié le ... :   /  /    
Description .. : Fonction indiquant la valeur de CPT_DATEVALID pour une 
Suite ........ : déclaration donnée.
Mots clefs ... : 
*****************************************************************}
function GetDateValidDeclaration(NumDeclaration : Integer) : TDateTime ;
var
   SQL : String;
   Q : TQuery;
begin
   result := 0;
   SQL := 'SELECT CPT_DATEVALID FROM PARAMEXTTVA WHERE CPT_CODE="' + IntToStr(NumDeclaration) + '"';
   try
     Q := OpenSQL(SQL,true);
     If Not Q.Eof then
     begin
       result := Q.Fields[0].AsDateTime;
     end;
   finally
     ferme(Q);
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/01/2007
Modifié le ... :   /  /    
Description .. : Fonction indiquant la valeur de CPT_NATURE pour une 
Suite ........ : déclaration donnée.
Mots clefs ... : 
*****************************************************************}
function GetNatureDeclaration(      NumDeclaration : Integer)   : string;
var
   SQL : String;
   Q : TQuery;
begin
   result := '';
   SQL := 'SELECT CPT_NATURE FROM PARAMEXTTVA WHERE CPT_CODE="' + IntToStr(NumDeclaration) + '"';
   try
     Q := OpenSQL(SQL,true);
     If Not Q.Eof then
     begin
       result := Q.Fields[0].AsString;
     end;
   finally
     ferme(Q);
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/01/2007
Modifié le ... :   /  /    
Description .. : Fonction indiquant la valeur de CPT_COLLECTIF pour
Suite ........ : une déclaration donnée.
Mots clefs ... : 
*****************************************************************}
function GetTypeColDeclaration(     NumDeclaration : Integer)   : string;
var
   SQL : String;
   Q : TQuery;
begin
   result := '';
   SQL := 'SELECT CPT_COLLECTIF FROM PARAMEXTTVA WHERE CPT_CODE="' + IntToStr(NumDeclaration) + '"';
   try
     Q := OpenSQL(SQL,true);
     If Not Q.Eof then
     begin
       result := Q.Fields[0].AsString;
     end;
   finally
     ferme(Q);
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/01/2007
Modifié le ... :   /  /    
Description .. : Fonction indiquant la valeur de CPT_ANNEE pour une 
Suite ........ : déclaration donnée.
Mots clefs ... :
*****************************************************************}
function GetAnneeDeclaration(   NumDeclaration : Integer)   : string;
var
   SQL : String;
   Q : TQuery;
begin
   result := '';
   SQL := 'SELECT CPT_ANNEE FROM PARAMEXTTVA WHERE CPT_CODE="' + IntToStr(NumDeclaration) + '"';
   try
     Q := OpenSQL(SQL,true);
     If Not Q.Eof then
     begin
       result := Q.Fields[0].AsString;
     end;
   finally
     ferme(Q);
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : Fonction indiquant la valeur de CPT_TRIMESTRE pour une 
Suite ........ : déclaration donnée.
Mots clefs ... : 
*****************************************************************}
function GetTrimDeclaration(   NumDeclaration : Integer)   : string;
var
   SQL : String;
   Q : TQuery;
begin
   result := '';
   SQL := 'SELECT CPT_TRIMESTRE FROM PARAMEXTTVA WHERE CPT_CODE="' + IntToStr(NumDeclaration) + '"';
   try
     Q := OpenSQL(SQL,true);
     If Not Q.Eof then
     begin
       result := Q.Fields[0].AsString;
     end;
   finally
     ferme(Q);
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/01/2007
Modifié le ... :   /  /    
Description .. : Fonction indiquant la valeur de CPT_REGIMETVA pour une 
Suite ........ : déclaration donnée.
Mots clefs ... : 
*****************************************************************}
function GetRegimeTVADeclaration(   NumDeclaration : Integer)   : string;
var
   SQL : String;
   Q : TQuery;
begin
   result := '';
   SQL := 'SELECT CPT_REGIMETVA FROM PARAMEXTTVA WHERE CPT_CODE="' + IntToStr(NumDeclaration) + '"';
   try
     Q := OpenSQL(SQL,true);
     If Not Q.Eof then
     begin
       result := Q.Fields[0].AsString;
     end;
   finally
     ferme(Q);
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/01/2007
Modifié le ... :   /  /    
Description .. : Fonction retournant sous un TDateTime la date de debut
Suite ........ : d'exercice de la déclaration passée en paramètre
Mots clefs ... :
*****************************************************************}
function GetDebutExercice(NumDeclaration : Integer) : TDateTime ;
var
   SQL : String;
   Q : TQuery;
begin
   SQL := 'SELECT CPT_DEBPERIODE FROM PARAMEXTTVA WHERE CPT_CODE="' + IntToStr(NumDeclaration) + '"';
   try
     Q := OpenSQL(SQL,true);
     If Not Q.Eof then
       result := Q.Fields[0].AsDateTime
     else
       Result := 0;
   finally
     ferme(Q);
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : Fonction retournant sous un TDateTime la date de
Suite ........ : fin d'exercice de la déclaration passée en paramètre
Mots clefs ... :
*****************************************************************}
function GetFinExercice(NumDeclaration : Integer) : TDateTime  ;
var
   SQL : String;
   Q : TQuery;
begin
   SQL := 'SELECT CPT_FINPERIODE FROM PARAMEXTTVA WHERE CPT_CODE="' + IntToStr(NumDeclaration) + '"';
   try
     Q := OpenSQL(SQL,true);
     If Not Q.Eof then
       result := Q.Fields[0].AsDateTime
     else
       Result := 0;
   finally
     ferme(Q);
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : Fonction retournant l'année la plus ancienne paramétrée
Suite ........ : dans la table EXERCICE.
Mots clefs ... :
*****************************************************************}
function GetMinAnneeExercice : integer ;
var
   SQL : String;
   Q : TQuery;
begin
   result := 1900;
   SQL := 'SELECT YEAR(MIN(EX_DATEDEBUT)) FROM EXERCICE';
   try
     Q := OpenSQL(SQL,true);
     If Not Q.Eof then
     begin
       result := Q.Fields[0].AsInteger;
     end;
   finally
     ferme(Q);
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : Fonction retournant l'année la plus récente paramétrée dans
Suite ........ : la table EXERCICE.
Mots clefs ... :
*****************************************************************}
function GetMaxAnneeExercice : integer ;
var
   SQL : String;
   Q : TQuery;
begin
   result := 2100;
   SQL := 'SELECT YEAR(MAX(EX_DATEFIN)) FROM EXERCICE';
   try
     Q := OpenSQL(SQL,true);
     If Not Q.Eof then
     begin
       result := Q.Fields[0].AsInteger;
     end;
   finally
     ferme(Q);
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/01/2007
Modifié le ... :   /  /    
Description .. : Fonction permettant de traiter les valeurs multiples des 
Suite ........ : combos.
Suite ........ : Prends une chaine de forme 123;456;789 et retourne 
Suite ........ : "123","456","789"
Mots clefs ... :
*****************************************************************}
function ReplaceChar(Combos : String) : string;
var
   i : integer;
begin
     result := '"';
     for i := 1 to (length(Combos) - 1)  do
     begin
          if Combos[i] = ';' then
             result := result + '","'
          else
             result := result + Combos[i];
     end;
     result := result + '"';
end;
 
{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 30/01/2007
Modifié le ... :   /  /    
Description .. : Fonction qui supprime dans une chaine l'ensemble des 
Suite ........ : occurences du caractère passé en parametre.
Mots clefs ... : 
*****************************************************************}
function DeleteChar(Source : String;Car : Char) : string;
var
  i,j : integer;
begin
  j := 0;
  for i := 1 to Length(Source) do
  begin
     if Source[i] = Car then continue;
     Inc(j);
     SetLength(Result,j);
     Result[j] := Source[i];
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/01/2007
Modifié le ... :   /  /    
Description .. : Procedure permettant de rajouter une chaine à la chaine 
Suite ........ : passée en  parametre à partir d'un certain caractère
Mots clefs ... : 
*****************************************************************}
procedure PutLigne(var Ligne : string; pos : integer; rajout : string);
var i,j,k    : integer;
begin
   k := pos + length(rajout) - 1;
   // On verifie que la Ligne n'est pas trop petite
   if Length(Ligne) < k then
   begin
     // C'est le cas on l'agrandi.
     Ligne := CompleteWithCar(Ligne,' ',k,'D');
   end;
   j := 1;
   for i := pos to k do
   begin
        Ligne[i] := Rajout[j];
        Inc(j);
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/01/2007
Modifié le ... :   /  /    
Description .. : Fonction permettant de retourner une chaine complete par
Suite ........ : un caractere sur une certaine longueur en partant de la 
Suite ........ : gauche ou la droite
Mots clefs ... : 
*****************************************************************}
function CompleteWithCar(Chaine : string ; Car : Char ; Long : integer; Sens : Char) : string;
var
   LongChaine : integer;
   i          : integer;
   Temp       : string;
begin 
   result := Chaine;
   LongChaine := Length(Chaine);
   if (Long <= LongChaine) then
      Exit
   else
       SetLength(Temp,Long - LongChaine);
   if (UpperCase(Sens) = 'G') then
   begin  
      { FQ 19783 BVE 13/03/2007 }
      if Chaine[1] = '-' then
      begin
         SetLength(Temp,Long - LongChaine + 1);
         Temp[1] := '-';
         for i := 2 to (Long - LongChaine + 1) do
            Temp[i] := Car;
         result := Temp;
         Temp := Chaine;
         Chaine := Copy(Chaine,2,LongChaine-1);
         result := Result + Chaine; 
      end
      else
      begin
         for i := 1 to (Long - LongChaine) do
            Temp[i] := Car;   
         result := Temp + Chaine;
      end;
      {END FQ 19783 BVE 13/03/2007 }
   end
   else if (UpperCase(Sens) = 'D') then
   begin
      for i := 1 to (Long - LongChaine) do
         Temp[i] := Car;
      result := Chaine + Temp;
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 29/01/2007
Modifié le ... :   /  /    
Description .. : Fonction permettant d'arrondir un double avec X chiffre
Suite ........ : apres la virgule
Mots clefs ... : 
*****************************************************************}
function ArrondirApresVirgule(Nombre : double; NbChiffre : integer) : Double;
var               
  virgule : boolean;
  temp    : string;
  res     : string;
  i,j     : integer;
begin
  virgule := false;
  temp := FloatToStr(Nombre);
  j := 0;
  for i := 1 to Length(temp) do
  begin
     if virgule then Inc(j);
     if j > NbChiffre then break;
     if (temp[i] = ',') or (temp[i] = '.') then
        virgule := true;
     setlength(res,i);
     res[i] := temp[i];
  end;
  result := StrToFloat(res);
end;

function LancerSaveDialog(var Repertoire : string) : Boolean;
var
  SD  : TSaveDialog;
begin
  Result := True;
  {Création et paramétrage de la boite de dialogue}
  SD := TSaveDialog.Create(nil);
  try
    SD.DefaultExt := 'TXT';
    SD.Filter := 'Fichier Texte (*.txt)|*.txt';
    SD.FilterIndex := 1;
    SD.Title := 'Declaration TVA';
    Result := SD.Execute;
    Repertoire := SD.FileName;
  finally
    if Assigned(SD) then FreeAndNil(SD);
  end;
end;
end.




