//-------------------------------------------------------
//--- Auteur : CATALA David
//--- Objet  : Librairie déversement du plan comptable
//-------------------------------------------------------
unit ULibCreerProduction;

interface

uses UTob,ParamSoc,HCtrls,SysUtils,Hent1,
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     HStatus,Ent1,UtobDebug;

function EviterDoublon(var ValChamp: variant; tab, champ: string; LeType: TFichierBase):Boolean;
function CompleterChaine ( St : String ; LgCompte : integer ):string;
procedure CreerSectionAttente (NomBase : string='');

procedure InitialiserTableEtablissement (NomBase : string);
procedure InitialiserTableEtablissementComplement (NomBase : String; Etablissement : String; Libelle : String; Localisation: string);
procedure InitialiserMonnaieDossier (NomBase : string );

procedure LoadStandardCompta (NumStd : integer; TableDos,TableStd : string; WhereDos:string=''; WhereStd : string=''; LgCpteGen : integer=-1; LgCpteAux : integer = -1);
procedure LoadStandardMaj(NumStd: integer; TableDos, TableStd, cle, condition: string ; RazFiltres : Boolean = false; LgCpteGen : integer=-1; LgCpteAux : integer = -1);

procedure DeverserParamSocRef (NomBase : String; NumPlan: integer; var LgCpteGen : integer; var LgCpteAux : integer );
procedure DeverserStandardDansDossier (NomBase : String; NumStandard : integer);

implementation

//-------------------------------------------
//--- Nom   : EviterDoublon
//--- Objet : Détermine les cas en doublon
//-------------------------------------------
function EviterDoublon(var ValChamp: variant; tab, champ: string; LeType: TFichierBase): Boolean;
var Val: string;
    OK: Boolean;
begin
  OK := TRUE;
  Val := VarToStr(ValChamp);
  if Length(Val) > VH^.Cpta[fbGene].Lg then
  begin
    Val := copy(ValChamp, 1, VH^.Cpta[LeType].Lg);
    if not Presence(tab, champ, Val) then
      ValChamp := Val
    else
      OK := FALSE;
  end;
  Result := OK;
end;

//-----------------------------------------------
//--- Nom   : CompleterChaine
//--- Objet : Complete une chaine de caractère
//-----------------------------------------------
function CompleterChaine (St : String ; LgCompte : integer ) : string ;
var ll,i : Integer ;
    Bourre  : Char ;
begin
  Bourre:='0';
  Result:=St;ll:=Length(Result);
  If ll<LgCompte then for i:=ll+1 to LgCompte do Result:=Result+Bourre;
end;

//----------------------------------
//--- Nom   : CreerSectionAttente
//--- Objet :
//----------------------------------
procedure CreerSectionAttente (NomBase : string='');
var stTableAxe, stTableSection, stSection : string;
    i: integer;
begin
  if NomBase='' then stTableAxe := 'AXE' else stTableAxe := NomBase+'.dbo.AXE';
  if NomBase='' then stTableSection := 'SECTION' else stTableAxe := NomBase+'.dbo.SECTION';
  for i := 1 to 5 do
    if VH^.Cpta[AxeToFb('A' + IntToStr(i))].Attente <> '' then
    begin
      stSection := BourreLaDonc(VH^.Cpta[AxeToFb('A' + IntToStr(i))].Attente,
        AxeToFb('A' + IntToStr(i)));
      ExecuteSQL('Update '+stTableAxe+' set X_SECTIONATTENTE="' + stSection +
        '" Where X_AXE="A' + IntToStr(i) + '"');
      if not ExisteSQL('SELECT * FROM '+stTableSection+' Where S_SECTION="' +
        VH^.Cpta[AxeToFb('A' + IntToStr(i))].Attente + '"') then
        ExecuteSQL('INSERT INTO '+stTableSection+' (S_SECTION,S_LIBELLE,S_ABREGE,S_SENS,S_SOLDEPROGRESSIF,S_AXE)'
          +
          ' VALUES ("' + stSection + '","' +
            TraduireMemoire('Section d''attente') +
          '","' + Copy(TraduireMemoire('Section d''attente'),1,17) + '","M","X","A' +
            IntToStr(i) + '")');
    end;
end;

//--------------------------------------------
//--- Nom   : InitialiserTableEtablissement
//--- Objet : Initialse la table ETABLISS
//--------------------------------------------
procedure InitialiserTableEtablissement (NomBase : string);
var ChTableEtablissement : String;
    ChTableSociete       : string;
    TobEtablissement     : Tob;
    TobSociete           : Tob;
begin
 //--- Détermine le nom des tables
 ChTableEtablissement:='ETABLISS';
{ ChTableSociete:='SOCIETE';
 if (NomBase<>'') then
  begin
   ChTableEtablissement:=NomBase+'.dbo.'+ChTableEtablissement;
   ChTableSociete:=NomBase+'.dbo.'+ChTableSociete;
  end;}

 TobEtablissement:=Tob.Create (ChTableEtablissement,Nil,-1);
 TobEtablissement.LoadDB;

// if (TobEtablissement.Detail.Count=0) then => non
//  begin
   TobSociete:=Tob.Create (ChTableSociete,Nil,-1);
   TobSociete.LoadDetailFromSQL('SELECT * FROM '+ChTableSociete);

   if (TobSociete.Detail.count>0) then
    begin
     TobEtablissement.PutValue ('ET_ETABLISSEMENT',TobSociete.Detail[0].GetValue('SO_SOCIETE'));
     TobEtablissement.PutValue ('ET_SOCIETE',TobSociete.Detail[0].GetValue('SO_SOCIETE'));
     TobEtablissement.PutValue ('ET_LIBELLE',TobSociete.Detail[0].GetValue('SO_LIBELLE'));
     TobEtablissement.PutValue ('ET_ABREGE',Copy (TobSociete.Detail[0].GetValue('SO_LIBELLE'),1,17));
     TobEtablissement.PutValue ('ET_ADRESSE1',TobSociete.Detail[0].GetValue('SO_ADRESSE1'));
     TobEtablissement.PutValue ('ET_ADRESSE2',TobSociete.Detail[0].GetValue('SO_ADRESSE2'));
     TobEtablissement.PutValue ('ET_ADRESSE3',TobSociete.Detail[0].GetValue('SO_ADRESSE3'));
     TobEtablissement.PutValue ('ET_CODEPOSTAL',TobSociete.Detail[0].GetValue('SO_CODEPOSTAL'));
     TobEtablissement.PutValue ('ET_VILLE',TobSociete.Detail[0].GetValue('SO_VILLE'));
     TobEtablissement.PutValue ('ET_PAYS',TobSociete.Detail[0].GetValue('SO_PAYS'));
     TobEtablissement.PutValue ('ET_TELEPHONE',TobSociete.Detail[0].GetValue('SO_TELEPHONE'));
     TobEtablissement.PutValue ('ET_FAX',TobSociete.Detail[0].GetValue('SO_FAX'));
     TobEtablissement.PutValue ('ET_TELEX',TobSociete.Detail[0].GetValue('SO_TELEX'));
     TobEtablissement.PutValue ('ET_SIRET',TobSociete.Detail[0].GetValue('SO_SIRET'));
     TobEtablissement.PutValue ('ET_APE',TobSociete.Detail[0].GetValue('SO_APE'));
     TobEtablissement.PutValue ('ET_JURIDIQUE',TobSociete.Detail[0].GetValue('SO_NATUREJURIDIQUE'));
     //ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="'+TobSociete.Detail[0].GetValue('SO_SOCIETE')+'" WHERE SOC_NOM="SO_ETABLISDEFAUT"');
     TobEtablissement.ChargeCle1;
     TobEtablissement.InsertDB (Nil);
    end;
   TobSociete.Free;
//  end;
 TobEtablissement.Free;
end;

//------------------------------------------------------
//--- Nom   : InitialiserTableEtablissementComplement
//--- Objet : Initialise la table ETABCOMPL
//------------------------------------------------------
procedure InitialiserTableEtablissementComplement (NomBase : String; Etablissement : String; Libelle : String; Localisation: string);
var ChTableEtablissementCompl : String;
    TobEtablissementCompl     : TOB;
    DateFermeture             : TDateTime;
    ANow,MNow,DNow            : Word;
begin
 //--- Détermine le nom de la table
 ChTableEtablissementCompl:='ETABCOMPL';
 if (NomBase<>'') then
  ChTableEtablissementCompl:=NomBase+'.dbo.'+ChTableEtablissementCompl;

 //--- Calcul DateFermetue
 DecodeDate (DATE, ANow, MNow, DNow);
 DateFermeture:=EncodeDate(ANow, 05, 31);
 if (DateFermeture<=Now) then
  DateFermeture := EncodeDate(ANow + 1, 05, 31);

 TobEtablissementCompl:=Tob.Create (ChTableEtablissementCompl,Nil,-1);
 TobEtablissementCompl.PutValue ('ETB_CONGESPAYES','-');
 TobEtablissementCompl.PutValue ('ETB_DATECLOTURECPN',DateFermeture);
 TobEtablissementCompl.PutValue ('ETB_JOURHEURE',(GetParamSocSecur('SO_PGJOURHEURE', 'HEU')));
 TobEtablissementCompl.PutValue ('ETB_CONGESPAYES', '-');
 TobEtablissementCompl.PutValue ('ETB_CODESECTION', '1');
 TobEtablissementCompl.PutValue('ETB_BCJOURPAIEMENT', 0);
 TobEtablissementCompl.PutValue('ETB_JEDTDU', 0);
 TobEtablissementCompl.PutValue('ETB_JEDTAU', 0);
 TobEtablissementCompl.PutValue('ETB_PRUDHCOLL', '1');
 TobEtablissementCompl.PutValue('ETB_PRUDHSECT', '4');
 TobEtablissementCompl.PutValue('ETB_PRUDHVOTE', '1');
 TobEtablissementCompl.PutValue('ETB_ISLICSPEC', '-');
 TobEtablissementCompl.PutValue('ETB_ISOCCAS', '-');
 TobEtablissementCompl.PutValue('ETB_ISLABELP', '-');
 TobEtablissementCompl.PutValue('ETB_ETABLISSEMENT', etablissement);
 TobEtablissementCompl.PutValue('ETB_LIBELLE', LIBELLE);
 TobEtablissementCompl.PutValue('ETB_NBJOUTRAV', 0);
 TobEtablissementCompl.PutValue('ETB_NBREACQUISCP', 0);
 TobEtablissementCompl.PutValue('ETB_NBACQUISCP', '');
 TobEtablissementCompl.PutValue('ETB_NBRECPSUPP', 0);
 TobEtablissementCompl.PutValue('ETB_TYPDATANC', '1');
 TobEtablissementCompl.PutValue('ETB_DATEACQCPANC', '');
 TobEtablissementCompl.PutValue('ETB_VALORINDEMCP', '');
 TobEtablissementCompl.PutValue('ETB_RELIQUAT', '');
 TobEtablissementCompl.PutValue('ETB_PROFILCGE', '');
 TobEtablissementCompl.PutValue('ETB_BASANCCP', '');
 TobEtablissementCompl.PutValue('ETB_VALANCCP', '');
 TobEtablissementCompl.PutValue('ETB_PERIODECP', 0);
 TobEtablissementCompl.PutValue('ETB_1ERREPOSH', '');
 TobEtablissementCompl.PutValue('ETB_2EMEREPOSH', '');
 TobEtablissementCompl.PutValue('ETB_MVALOMS', '');
 TobEtablissementCompl.PutValue('ETB_VALODXMN', 0);
 TobEtablissementCompl.PutValue('ETB_TYPDATANC', '');
 TobEtablissementCompl.PutValue('ETB_MVALOMS', '');
 TobEtablissementCompl.PutValue('ETB_PCTFRAISPROF', 0);
 TobEtablissementCompl.PutValue('ETB_HORAIREETABL', 0);
 TobEtablissementCompl.PutValue('ETB_DATEVALTRANS', Idate1900);
 TobEtablissementCompl.PutValue('ETB_SEUILTEMPSPAR', 0);
 TobEtablissementCompl.PutValue('ETB_PRORATATVA', 0);
 TobEtablissementCompl.PutValue('ETB_JOURPAIEMENT', 0);
 TobEtablissementCompl.PutValue('ETB_EDITBULCP', '');
 TobEtablissementCompl.PutValue('ETB_REGIMEALSACE', '-');
 TobEtablissementCompl.PutValue('ETB_MEDTRAV', -1);
 TobEtablissementCompl.PutValue('ETB_CODEDDTEFP', -1);
 TobEtablissementCompl.PutValue('ETB_SUBROGATION','-');

 //--- Enregistrement en base
 TobEtablissementCompl.InsertOrUpdateDB (FALSE);
 TobEtablissementCompl.Free;
end;

//----------------------------------------------
//--- Nom   : InitialiserMonnaieDossier
//--- Objet : Intialise la monnaie du dossier
//----------------------------------------------
procedure InitialiserMonnaieDossier (NomBase : string );
var ChSql : String;
    RSql  : TQuery;
begin
  //--- Mise à jour du taux par rapport à l'euro
  ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="1" WHERE SOC_NOM="SO_TAUXEURO"');
  //--- Mise à jour de l'enregistrement devise Euro
  RSql:=OpenSQL ('SELECT * FROM '+NomBase+'.dbo.DEVISE WHERE D_DEVISE="EUR"',True);
  try
   if not Rsql.Eof then
    begin
     ChSql:='UPDATE '+NomBase+'.dbo.DEVISE SET D_FERME="-", D_DECIMALE=2, D_QUOTITE=1, D_MONNAIEIN="-", D_FONGIBLE="-", D_PARITEEURO=1';
     ExecuteSql (ChSql);
     end
    else
     begin
      ChSQl:='INSERT INTO DEVISE (D_DEVISE,D_LIBELLE,D_SYMBOLE,D_FERME,D_DECIMALE,D_QUOTITE,D_SOCIETE,D_MONNAIEIN,D_FONGILE,D_PARITEEURO)';
      ChSql:=ChSql+' VALUES ("EUR","Euro","€","-",2,1,"'+V_PGI.CodeSociete+'", "-","-",1)';
      ExecuteSql (ChSql);
     end;
  finally
   Ferme (Rsql);
  end;
  ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="EUR" WHERE SOC_NOM="SO_DEVISEPRINC"');
  //--- Suppression des monnaie 'In'
  ExecuteSQL ('DELETE FROM '+NomBase+'.dbo.DEVISE WHERE D_MONNAIEIN="X" AND D_FONGIBLE="-"');
end;

//------------------------------------------------------
//--- Nom   : LoadStandardCompta
//--- Objet :
//------------------------------------------------------
procedure LoadStandardCompta (NumStd : integer; TableDos,TableStd : string; WhereDos:string=''; WhereStd : string=''; LgCpteGen : integer=-1; LgCpteAux : integer = -1);
var PrefDos,PrefStd,Suffixe,ChampDos : String;
    RSql                                   : TQuery;
    TobDossier, TobTiersComplementaire     : Tob;
    OkMaj                                  : Boolean;
    ValChamp                               : variant;
    i                                      : integer;
begin
 if LgCpteGen=-1 then
  LgCpteGen := VH^.Cpta[fbGene].Lg;
 if LgCpteAux=-1 then
  LgCpteAux := VH^.Cpta[fbAux].Lg;

 PrefDos := TableToPrefixe(TableDos);
 PrefStd := TableToPrefixe(TableStd);

 RSql:=OpenSQL ('SELECT * FROM '+TableStd+' WHERE '+PrefStd+'_NUMPLAN='+IntToStr(NumStd) + WhereStd,True);
 if not RSql.Eof then
  begin
{$IFDEF EAGLCLIENT}
   InitMove(100,'Chargement Table : '+TableDos);
{$ELSE}
   InitMove(QCount(RSql),'Chargement Table : '+TableDos);
{$ENDIF}

   TobDossier:=Tob.Create (TableDos,Nil,-1);
   TobDossier.ClearDetail;
   TobDossier.LoadDetailFromSQL('SELECT * FROM '+TableDos+WhereDos);

   if (TobDossier.Detail.Count=0) then
    begin
     while not (RSql.Eof) do
      begin
       OkMaj:=True;

       //--- Traitement des tiers complementaires
       if (Pos('TIERSREF',TableStd)>0) then
        begin
         TobTiersComplementaire:=Tob.Create ('TIERSCOMPL',Nil,-1);
         TobTiersComplementaire.ClearDetail;
         TobTiersComplementaire.LoadDetailFromSQL ('SELECT * FROM TIERSCOMPL WHERE YTC_AUXILIAIRE="'+CompleterChaine(RSql.FindField('TRR_AUXILIAIRE').AsString, LgCpteAux)+'"');

         if (TobTiersComplementaire.Detail.Count=0) then
          begin
           TobTiersComplementaire.PutValue ('YTC_AUXILIAIRE',CompleterChaine(RSql.FindField('TRR_AUXILIAIRE').AsString, LgCpteAux));
           TobTiersComplementaire.PutValue ('YTC_TIERS',RSql.FindField('TRR_TIERS').AsString);
           TobTiersComplementaire.PutValue ('YTC_SCHEMAGEN',RSql.FindField('TRR_SCHEMAGEN').AsString);
           TobTiersComplementaire.PutValue ('YTC_ACCELERATEUR',RSql.FindField('TRR_ACCELERATEUR').AsString);
          end;

         TobTiersComplementaire.InsertDB (Nil);
         TobTiersComplementaire.Free;
        end;

       //--- Traitement des standards
       for i:=0 to RSql.FieldCount-1 do
        begin
         Suffixe:=ExtractSuffixe (RSql.Fields[i].FieldName);
         // Gestion des exceptions à la règle ici
         if ((PrefStd='PR') and (Suffixe='COMPTE')) then ChampDos := 'G_GENERAL'
         else if ((PrefStd='PR') and (Suffixe='REPORTDETAIL')) then continue
         else if Suffixe = 'PREDEFINI' then continue
         else if Suffixe = 'NUMPLAN' then continue
         // GCO - 01/06/2004 - Gestion des champs Accelerateur
         else if (TableStd = 'TIERSREF') and ((Suffixe = 'SCHEMAGEN') or (Suffixe = 'ACCELERATEUR')) then Continue
         else if (TableStd = 'JALREF') and (Suffixe = 'ACCELERATEUR') then Continue
         // Fin de la gestion des exceptions
         else ChampDos:=PrefDos+'_'+Suffixe;
         // calcul de la valeur du champ
         if (ChampDos='G_GENERAL') then
          begin
           ValChamp := CompleterChaine(RSql.FindField('GER_GENERAL').AsString,LgCpteGen);
           OkMaj := EviterDoublon (ValChamp,'GENERAUX','G_GENERAL',fbGene);
          end
         else
          if (ChampDos='T_AUXILIAIRE') then
           begin
            ValChamp:=CompleterChaine(RSql.FindField('TRR_AUXILIAIRE').AsString,LgCpteAux);
            Okmaj:=EviterDoublon (ValChamp,'TIERS','T_AUXILIAIRE',fbAux);
           end
          else if ChampDos = 'T_TIERS' then ValChamp := CompleterChaine(RSql.FindField('TRR_TIERS').AsString,LgCpteAux)
          else if ChampDos = 'T_COLLECTIF' then ValChamp := CompleterChaine(RSql.FindField('TRR_COLLECTIF').AsString,LgCpteGen)
          else ValChamp:=RSql.FindField(PrefStd+'_'+Suffixe).AsVariant;

         TobDossier.PutValue (ChampDos,ValChamp);
        end;

        if OkMaj then
         TobDossier.InsertDB (Nil)
        else
         TobDossier:=Nil;

        MoveCur(False);
        RSql.Next;
      end;
    end;

   TobDossier.Free;
   FiniMove;
  end;
 Ferme (RSql);
end;

//------------------------------
//--- Nom   : LoadStandardMaj
//--- Objet :
//------------------------------
procedure LoadStandardMaj(NumStd: integer; TableDos, TableStd, cle, condition: string ; RazFiltres : Boolean = false; LgCpteGen : integer=-1; LgCpteAux : integer = -1);
begin
end;

//----------------------------------
//--- Nom   : DeverserParamSocRef
//--- Objet :
//----------------------------------
procedure DeverserParamSocRef (NomBase : String; NumPlan: integer; var LgCpteGen : integer; var LgCpteAux : integer );
var ChSql, StVal : String;
    RSql         : TQuery;
begin
 LgCpteGen := 6;
 LgCpteAux := 6;

 //--- Longueur des comptes généraux
 ChSql:='SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN= '+IntToStr(NumPlan)+' AND PRR_SOCNOM="SO_LGCPTEGEN"';
 RSql:=OpenSQL(ChSql,True);
 if (RSql.Eof) then
  ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="6" WHERE SOC_NOM="SO_LGCPTEGEN"')
 else
  begin
   ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="'+RSql.FindField('PRR_SOCDATA').AsString+'" WHERE SOC_NOM="SO_LGCPTEGEN"');
   LgCpteGen:=RSql.FindField('PRR_SOCDATA').AsInteger;
  end;
 ferme (RSql);

 //--- Longueur des comptes auxiliaires
 ChSql:='SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN= '+IntToStr(NumPlan)+' AND PRR_SOCNOM="SO_LGCPTEAUX"';
 RSql:=OpenSql (ChSql,True);
 if (RSql.Eof) then
  ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="6" WHERE SOC_NOM="SO_LGCPTEAUX"')
 else
  begin
   ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="'+RSql.FindField('PRR_SOCDATA').AsString+'" WHERE SOC_NOM="SO_LGCPTEAUX"');
   LgCpteAux:=RSql.FindField('PRR_SOCDATA').AsInteger;
  end;
 ferme (RSql);

 //--- Paramètres sociétés
 ChSql:='SELECT * FROM PARSOCREF WHERE PRR_NUMPLAN= '+IntToStr(NumPlan)+' ORDER BY PRR_SOCNOM';
 RSql:=OpenSql (ChSql,True);

 InitMove(RSql.RecordCount, 'Chargement des paramètres société');
 while not (RSql.Eof) do
  begin
    StVal:=RSql.FindField('PRR_SOCDATA').AsString;
    if (RSql.FindField('PRR_COMPTE').AsString = 'G') then
     StVal := CompleterChaine(stVal, LgCpteGen)
    else
     if (RSql.FindField('PRR_COMPTE').AsString = 'T') then
      StVal := CompleterChaine(stVal, LgCpteAux);
    ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="'+StVal+'" WHERE SOC_NOM="'+RSql.FindField('PRR_SOCNOM').AsString+'"');
    RSql.Next;
    MoveCur(False);
  end;

  Ferme(RSql);
  FiniMove;
end;

//------------------------------------------
//--- Nom   : DeverserStandardDansDossier
//--- Objet :
//------------------------------------------
procedure DeverserStandardDansDossier (NomBase : String; NumStandard : integer);
var LgCpteGen, LgCpteAux : integer;
begin
 InitialiserTableEtablissement (NomBase);
 DeverserParamSocRef(NomBase,NumStandard,LgCpteGen,LgCpteAux);
 ExecuteSQL ('UPDATE '+NomBase+'.dbo.PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM="SO_TENUEEURO"');
 InitialiserMonnaieDossier (NomBase );

 LoadStandardCompta(NumStandard, NomBase+'.dbo.GENERAUX', 'GENERAUXREF','','',LgCpteGen,LgCpteAux);
 LoadStandardCompta(NumStandard, NomBase+'.dbo.JOURNAL', 'JALREF','','',LgCpteGen,LgCpteAux);
 LoadStandardCompta(NumStandard, NomBase+'.dbo.GUIDE', 'GUIDEREF','','',LgCpteGen,LgCpteAux);
 LoadStandardCompta(NumStandard, NomBase+'.dbo.ECRGUI', 'ECRGUIREF','','',LgCpteGen,LgCpteAux);
 LoadStandardCompta(NumStandard, NomBase+'.dbo.ANAGUI', 'ANAGUIREF','','',LgCpteGen,LgCpteAux);
 LoadStandardCompta(NumStandard, NomBase+'.dbo.TIERS', 'TIERSREF','','',LgCpteGen,LgCpteAux);

 if ExisteSQL('SELECT * FROM MODEPAIEREF WHERE MPR_NUMPLAN='+IntToStr(NumStandard)) then
  begin
    ExecuteSQL('DELETE FROM '+NomBase+'.dbo.MODEPAIE');
    LoadStandardCompta(NumStandard, NomBase+'.dbo.MODEPAIE', 'MODEPAIEREF','','',LgCpteGen,LgCpteAux);
  end;

 if ExisteSQL('SELECT * FROM MODEREGLREF WHERE MRR_NUMPLAN='+IntToStr(NumStandard)) then
  begin
    ExecuteSQL('DELETE FROM '+NomBase+'.dbo.MODEREGL');
    LoadStandardCompta(NumStandard, NomBase+'.dbo.MODEREGL', 'MODEREGLREF','','',LgCpteGen,LgCpteAux);
  end;

 LoadStandardCompta(NumStandard, NomBase+'.dbo.CORRESP', 'CORRESPREF','','',LgCpteGen,LgCpteAux);
 LoadStandardCompta(NumStandard, NomBase+'.dbo.RUPTURE', 'RUPTUREREF','','',LgCpteGen,LgCpteAux);
 LoadStandardCompta(NumStandard, NomBase+'.dbo.REFAUTO', 'REFAUTOREF','','',LgCpteGen,LgCpteAux);
 LoadStandardCompta(NumStandard, NomBase+'.dbo.NATCPTE', 'NATCPTEREF','','',LgCpteGen,LgCpteAux);
 LoadStandardCompta(NumStandard, NomBase+'.dbo.FILTRES', 'FILTRESREF','','',LgCpteGen,LgCpteAux);

 LoadStandardMaj(NumStandard, NomBase+'.dbo.LISTE', 'LISTEREF', 'LISTE', ' ',False,LgCpteGen,LgCpteAux);
 LoadStandardMaj(NumStandard, NomBase+'.dbo.AXE', 'AXEREF', 'AXE', ' ',False,LgCpteGen,LgCpteAux);
 LoadStandardCompta(NumStandard, NomBase+'.dbo.SECTION', 'SECTIONREF','','',LgCpteGen,LgCpteAux);

 LoadStandardMaj(NumStandard, NomBase+'.dbo.CHOIXCOD', 'CHOIXCODREF', 'TYPE;CODE', ' ',False,LgCpteGen,LgCpteAux);
 LoadStandardCompta(NumStandard, NomBase+'.dbo.VENTIL', 'VENTILREF','','',LgCpteGen,LgCpteAux);

 // Installation des coefficients dégressifs Immo
 //CreationSectionAttente ( NomBase );
 //InstalleLesCoefficientsDegressifs ( NomBase ) ;
end;


end.
