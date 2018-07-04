//------------------------------------------------------------------------------
//--- Auteur : CATALA David
//--- Objet  :
//------------------------------------------------------------------------------
unit ObligationFonction;

interface

uses Utob,Hctrls,Hent1,Sysutils,HmsgBox,classes,paramsoc,winprocs,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     galoutil,Dialogs,UAFO_REGLES,AFPlanningCst,uDownloadUpdate,UCegidUpdate,
     Ed_Tools,CAT_FichierTexte,CAT_LibrairieFonction;

//-------------------------------------------
//--- Déclaration des procedures/fonctions
//-------------------------------------------
function DonnerRepertoireTemp : String;
function DonnerAccesObligation (TypeObligation : String; Utilisateur : String) : Boolean;
function DonnerDebutSemaine (DateSelection : TDateTime) : TDateTime;
function DonnerDateDebutPeriode (DateSelection : TDateTime; TypePeriode : Integer): TDateTime;
function DonnerDateFinPeriode (DateSelection : TDateTime; TypePeriode : Integer): TDateTime;
function DonnerDateFinValidite (DateGeneration : TDateTime; TypePeriode : String; Alerte : Boolean) : TDateTime;

function DonnerListeObligation (DomaineObligation : String; DateSelection : TDateTime; TypePeriode : Integer) : Tob;
function DonnerListeObligationDossier (NumDossier : String; DomaineObligation : String; DateSelection : TDateTime; TypePeriode : Integer) : Tob;
function DonnerDetailObligation (CodeObligation : String; DateObligation : String; NumDossier : String) : Tob;
function DonnerLibelleTypeDeclaration (TypeDeclaration : String) : String;
function DonnerFichierTelecharger (UneTob : Tob) : Integer;

function ChargerParametreObligation (Mode : String) : Tob;
function TelechargerListeObligation (RepertoireDest : String; NomFichier : String) :Boolean;

function ImporterListeObligation (CheminFichierObligation : String; NomFichier : String; Mode : String) : Boolean;

function ChargerListeObligation (CheminFichier : String) : Tob;
procedure InitialiserTableModeleTache (UneTob : Tob);
procedure CalculerListeObligation (DateDebutPeriode : TDateTime;DateFinPeriode : TDateTime);
procedure GenererObligation (Datedeb, datefin : Tdatetime;var  pTobMere :tob);

implementation

//-----------------------------------------------------------------------------
//--- Nom   : DonnerRepertoireTemp
//--- Objet :
//-----------------------------------------------------------------------------
function DonnerRepertoireTemp : String;
var Buffer: array[0..1023] of Char;
    ChRepertoire : String;
begin
 GetTempPath(Sizeof(Buffer)-1,Buffer);
 SetString(ChRepertoire, Buffer, StrLen(Buffer));
 Result:=ChRepertoire;
end;

//-----------------------------------------------------------------------------
//--- Nom   : DonnerAccesObligation
//--- Objet :
//-----------------------------------------------------------------------------
function DonnerAccesObligation (TypeObligation : String; Utilisateur : String):Boolean;
var ChSql    : String;
    RSql     : TQuery;
    ChChaine : String;
    TypeLu   : String;
begin
 Result:=False;
 ChSql:='SELECT CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="DPO" AND CC_CODE="'+Utilisateur+'"';
 RSql:=OpenSql (ChSql,True);
 if (not Rsql.Eof) then
  begin
   ChChaine:= RSql.FindField ('CC_LIBELLE').AsString;
   TypeLu:=(Trim(ReadTokenSt(ChChaine)));
   while (TypeLu<>'') do
    begin
     if (TypeLu=TypeObligation) then Result:=True;
     TypeLu:=(Trim(ReadTokenSt(ChChaine)));
    end;
  end;
 Ferme (RSql);
end;

//-----------------------------------------------------------------------------
//--- Nom   : DonnerDebutSemaine
//--- Objet :
//-----------------------------------------------------------------------------
function DonnerDebutSemaine (DateSelection : TDateTime) : TDateTime;
var JourSemaine : Integer;
begin
 JourSemaine:=DayOfWeek (DateSelection);
 if (JourSemaine=1) then
  Result:=DateSelection-6
 else
  Result:=DateSelection-(JourSemaine-2);
end;

//-----------------------------------------------------------------------------
//--- Nom   : DonnerDateDebutPeriode
//--- Objet :
//-----------------------------------------------------------------------------
function DonnerDateDebutPeriode (DateSelection : TDateTime; TypePeriode : Integer) : TDateTime;
begin
 Result:=iDate1900;
 if (TypePeriode=0) then Result:=DateSelection;
 if (TypePeriode=1) then Result:=DonnerDebutSemaine (DateSelection);
 if (TypePeriode=2) then Result:=DebutDemois (DateSelection);
 if (TypePeriode=3) then Result:=DonnerDebutSemaine (DateSelection);
end;

//-----------------------------------------------------------------------------
//--- Nom   : DonnerDateFinPeriode
//--- Objet :
//-----------------------------------------------------------------------------
function DonnerDateFinPeriode (DateSelection : TDateTime; TypePeriode : Integer) : TDateTime;
begin
 Result:=iDate1900;
 if (TypePeriode=0) then Result:=DateSelection;
 if (TypePeriode=1) then Result:=DonnerDebutSemaine (DateSelection)+6;
 if (TypePeriode=2) then Result:=FinDemois (DateSelection);
 if (TypePeriode=3) then Result:=PlusDate (DonnerDebutSemaine (DateSelection),4,'S');
end;

//-----------------------------------------------------------------------------
//--- Nom   : DonnerDateFinValidite
//--- Objet :
//-----------------------------------------------------------------------------
function DonnerDateFinValidite (DateGeneration : TDateTime; TypePeriode : String; Alerte : Boolean) : TDateTime;
var Duree : Integer;
begin
 Result:=iDate1900;
 if Alerte then Duree:=2 else Duree:=1;
 if (TypePeriode='ANN') then Result:=PlusDate (DateGeneration,Duree,'A');
 if (TypePeriode='SEM') then Result:=PlusDate (DateGeneration,Duree*6,'M');
 if (TypePeriode='TRI') then Result:=PlusDate (DateGeneration,Duree*3,'M');
 if (TypePeriode='BIM') then Result:=PlusDate (DateGeneration,Duree*2,'M');
 if (TypePeriode='MEN') then Result:=PlusDate (DateGeneration,Duree,'M');
 if (TypePeriode='HEB') then Result:=PlusDate (DateGeneration,Duree,'S');
 if (TypePeriode='JOU') then Result:=PlusDate (DateGeneration,Duree*7,'J');
end;

//------------------------------------------------------------------------------
//--- Nom   : DonnerListeObligation
//--- Objet : Donne la liste des obligations pour une periode donnee et
//---         conserve uniquement les obligations correspondant au domaine
//------------------------------------------------------------------------------
function DonnerListeObligation (DomaineObligation : String; DateSelection : TDateTime; TypePeriode : Integer) : Tob;
var DateDebutPeriode : TDateTime;
    DateFinPeriode   : TDateTime;
    UneTob           : Tob;
    ChSql            : String;
begin
 DateDebutPeriode:=DonnerDateDebutPeriode (DateSelection,TypePeriode);
 DateFinPeriode:=DonnerDateFinPeriode (DateSelection,TypePeriode);

 ChSql:='SELECT AFM_MODELETACHE, AFM_LIBELLETACHE1, AFM_LIBELLETACHE2, AFM_FAMILLETACHE, AFM_PLANNINGPERIOD, AFM_MOISJOURFIXE, AFM_COMMENTAIRE, '+
        'AFM_MOISFIXE, AFM_DATEANNUELLE, AFM_DESCRIPTIF, AFM_DATEDEBPERIOD, AFM_DATEFINPERIOD, AFM_CHARLIBRE3, DPT_DATEOBLIGATION '+
        'FROM AFMODELETACHE,DPOBLIGATIONTMP WHERE '+
        'AFM_MODELETACHE=DPT_CODEOBLIGATION AND '+
        'DPT_DATEOBLIGATION>="'+UsDateTime (DateDebutPeriode)+'" AND '+
        'DPT_DATEOBLIGATION<="'+UsDateTime (DateFinPeriode)+'" ';
        if (DomaineObligation<>'') then ChSql:=ChSql+'AND AFM_FAMILLETACHE="'+DomaineObligation+'"';

 UneTob:=Tob.Create ('',nil,0);
 UneTob.LoadDetailFromSql (ChSql);
 UneTob.Detail.Sort ('AFM_MOISJOURFIXE;AFM_LIBELLETACHE1');
 Result:=UneTob;
end;

//----------------------------------------------------------------------------------
//--- Nom   : DonnerListeObligationDossier
//--- Objet : Donne la liste des obligations d'un dossier pour une periode donnee
//----------------------------------------------------------------------------------
function DonnerListeObligationDossier (NumDossier : String; DomaineObligation : String; DateSelection : TDateTime; TypePeriode : Integer) : Tob;
var DateDebutPeriode      : TDateTime;
    DateFinPeriode        : TDateTime;
    UneTob,UneTobEnreg    : Tob;
    ChCodeObligation      : String;
    ChDateObligation      : String;
    ChSql                 : String;
    RSql                  : TQuery;
    RSqlObligationDossier : TQuery;
    RSqlObligationRealise : TQuery;
begin
 UneTob:=Tob.Create ('',nil,-1);
 DateDebutPeriode:=DonnerDateDebutPeriode (DateSelection,TypePeriode);
 DateFinPeriode:=DonnerDateFinPeriode (DateSelection,TypePeriode);

 //--- Récupère la liste des obligations + régle d'application
 ChSql:='SELECT DPT_CODEOBLIGATION, DPT_DATEOBLIGATION, AFM_DESCRIPTIF FROM DPOBLIGATIONTMP,AFMODELETACHE'+
        ' WHERE AFM_MODELETACHE=DPT_CODEOBLIGATION'+
        ' AND AFM_FAMILLETACHE="'+DomaineObligation+'"'+
        ' AND (DPT_DATEOBLIGATION>="'+UsDateTime(DateDebutPeriode)+'" AND DPT_DATEOBLIGATION<="'+UsDateTime(DateFinPeriode)+'")';
 RSql:=OpenSql (ChSql,True,-1,'',True);

 //--- Conserve uniquement les obligations correspondant à la régle d'application
 while (not RSql.Eof) do
  begin
   ChSql:='SELECT DOS_NODOSSIER FROM DPDOSSIEROBLIG'+
          ' WHERE DOS_NODOSSIER="'+NumDossier+'"'+
          ' AND '+RSql.FindField ('AFM_DESCRIPTIF').AsString;

   if (ExisteSQl (ChSql)) then
    begin
     ChCodeObligation:=RSql.FindField ('DPT_CODEOBLIGATION').AsString;
     ChDateObligation:=RSql.FindField ('DPT_DATEOBLIGATION').AsString;

     ChSql:='SELECT AFM_MODELETACHE, AFM_LIBELLETACHE1, AFM_LIBELLETACHE2, AFM_FAMILLETACHE, AFM_PLANNINGPERIOD, AFM_MOISJOURFIXE, AFM_COMMENTAIRE, '+
            'AFM_MOISFIXE, AFM_DATEANNUELLE, AFM_DESCRIPTIF, AFM_DATEDEBPERIOD, AFM_DATEFINPERIOD, AFM_CHARLIBRE3 FROM AFMODELETACHE'+
            ' WHERE AFM_MODELETACHE="'+ChCodeObligation+'"';

     RSqlObligationDossier:=OpenSql (ChSql,True);

     UneTobEnreg:=Tob.Create ('',UneTob,-1);
     UneTobEnreg.AddChampSupValeur ('AFM_MODELETACHE',ChCodeObligation);
     UneTobEnreg.AddChampSupValeur ('DOS_NODOSSIER',NumDossier);
     UneTobEnreg.AddChampSupValeur ('DPT_DATEOBLIGATION',StrToDate (ChDateObligation));

     UneTobEnreg.AddChampSupValeur ('AFM_LIBELLETACHE1',RSqlObligationDossier.FindField ('AFM_LIBELLETACHE1').AsString);
     UneTobEnreg.AddChampSupValeur ('AFM_LIBELLETACHE2',RSqlObligationDossier.FindField ('AFM_LIBELLETACHE2').AsString);
     UneTobEnreg.AddChampSupValeur ('AFM_FAMILLETACHE',RSqlObligationDossier.FindField ('AFM_FAMILLETACHE').AsString);
     UneTobEnreg.AddChampSupValeur ('AFM_PLANNINGPERIOD',RSqlObligationDossier.FindField ('AFM_PLANNINGPERIOD').AsString);
     UneTobEnreg.AddChampSupValeur ('AFM_CHARLIBRE3',RSqlObligationDossier.FindField ('AFM_CHARLIBRE3').AsString);
     UneTobEnreg.AddChampSupValeur ('AFM_COMMENTAIRE',RSqlObligationDossier.FindField ('AFM_COMMENTAIRE').AsString);

     UneTobEnreg.AddChampSupValeur ('AFM_MOISJOURFIXE',RSqlObligationDossier.FindField ('AFM_MOISJOURFIXE').AsString);

     //--- Périodicité annuel on prend une fourchette de 2 mois
     if (RSqlObligationDossier.FindField ('AFM_PLANNINGPERIOD').AsString='A') then
      ChSql:='SELECT DPO_UTILISATEUR, DPO_ETAT, DPO_COMMENTAIRE FROM DPOBLIGATIONREALISE WHERE DPO_CODEOBLIGATION="'+ChCodeObligation+'" AND DPO_NODOSSIER="'+NumDossier+'" AND DPO_DATEOBLIGATION>="'+UsDateTime (PlusMois (StrToDate (ChDateObligation),-2))+'" AND DPO_DATEOBLIGATION<="'+UsDateTime (StrToDate (ChDateObligation))+'"'
     else
      ChSql:='SELECT DPO_UTILISATEUR, DPO_ETAT, DPO_COMMENTAIRE FROM DPOBLIGATIONREALISE WHERE DPO_CODEOBLIGATION="'+ChCodeObligation+'" AND DPO_NODOSSIER="'+NumDossier+'" AND DPO_DATEOBLIGATION="'+UsDateTime (StrToDate (ChDateObligation))+'"';

     UneTobEnreg.AddChampSupValeur ('DPO_UTILISATEUR','');
     UneTobEnreg.AddChampSupValeur ('ISREALISE','');
     UneTobEnreg.AddChampSupValeur ('ISCOMMENTAIRE','');
     UneTobEnreg.AddChampSupValeur ('REALISE','');
     UneTobEnreg.AddChampSupValeur ('COMMENTAIRE','');

     if (ExisteSQL (ChSql)) then
      begin
       RSqlObligationRealise:=OpenSql (ChSql,True);
       UneTobEnreg.AddChampSupValeur ('DPO_UTILISATEUR',RSqlObligationRealise.FindField ('DPO_UTILISATEUR').AsString);
       UneTobEnreg.AddChampSupValeur ('COMMENTAIRE',RSqlObligationRealise.FindField ('DPO_COMMENTAIRE').AsString);
       Ferme (RSqlObligationRealise);
      end;

     if (ExisteSQL (ChSql+' AND DPO_COMMENTAIRE<>""')) then
      UneTobEnreg.AddChampSupValeur ('ISCOMMENTAIRE','#ICO#7');

     if (ExisteSQL (ChSql+' AND DPO_ETAT=1')) then
      begin
       UneTobEnreg.AddChampSupValeur ('ISREALISE','#ICO#3');
       UneTobEnreg.AddChampSupValeur ('REALISE','Réalisé');
      end
     else
      if (ExisteSQL (ChSql+' AND DPO_ETAT=2')) then
       begin
        UneTobEnreg.AddChampSupValeur ('ISREALISE','#ICO#12');
        UneTobEnreg.AddChampSupValeur ('REALISE','Réalisé client');
       end
      else
       if (ExisteSQL (ChSql+' AND DPO_ETAT=3')) then
        begin
         UneTobEnreg.AddChampSupValeur ('ISREALISE','#ICO#2');
         UneTobEnreg.AddChampSupValeur ('REALISE','Non applicable');
        end;
     Ferme (RSqlObligationDossier);
    end;
   RSql.next;
  end;

 Ferme (Rsql);
 Result:=UneTob;
end;

//-----------------------------------------------------------------------------
//--- Nom   : DonnerDetailObligation
//--- Objet :
//-----------------------------------------------------------------------------
function DonnerDetailObligation (CodeObligation : String; DateObligation : String; NumDossier : String) : Tob;
var ChSql  : String;
    UneTob : Tob;
    RSqlAfModeleTache       : TQuery;
    RSqlDpObligationRealise : TQuery;
begin
 UneTob:=Nil;
 ChSql:='SELECT AFM_LIBELLETACHE1, AFM_LIBELLETACHE2, AFM_FAMILLETACHE, AFM_PLANNINGPERIOD, AFM_MOISFIXE FROM AFMODELETACHE '+
        'WHERE AFM_MODELETACHE="'+CodeObligation+'"';
 RSqlAfModeleTache:=OpenSql (ChSql,True);

 if (not RSqlAfModeleTache.Eof) then
  begin
   if (RSqlAfModeleTache.FindField ('AFM_PLANNINGPERIOD').AsString='A') then
    ChSql:='SELECT DPO_DATEREALISE, DPO_UTILISATEUR, DPO_ETAT, DPO_COMMENTAIRE FROM DPOBLIGATIONREALISE '+
           'WHERE DPO_CODEOBLIGATION="'+CodeObligation+'" '+
           'AND DPO_NODOSSIER="'+NumDossier+'" '+
           'AND DPO_DATEOBLIGATION>="'+UsDateTime (PlusMois (StrToDate (DateObligation),-2))+'" '+
           'AND DPO_DATEOBLIGATION<="'+UsDateTime (StrToDate (DateObligation))+'"'
   else
    ChSql:='SELECT DPO_DATEREALISE, DPO_UTILISATEUR, DPO_ETAT, DPO_COMMENTAIRE FROM DPOBLIGATIONREALISE '+
           'WHERE DPO_CODEOBLIGATION="'+CodeObligation+'" '+
           'AND DPO_NODOSSIER="'+NumDossier+'" '+
           'AND DPO_DATEOBLIGATION="'+USDateTime (StrToDate (DateObligation))+'"';

   RSqlDpObligationRealise:=OpenSql (ChSql,True);

   if (not RSqlDpObligationRealise.Eof) then
    begin
     UneTob:=Tob.Create ('',nil,-1);
     UneTob.AddChampSupValeur ('AFM_MODELETACHE',CodeObligation);
     UneTob.AddChampSupValeur ('DOS_NODOSSIER',NumDossier);
     UneTob.AddChampSupValeur ('DPT_DATEOBLIGATION',StrToDate (DateObligation));

     UneTob.AddChampSupValeur ('AFM_LIBELLETACHE1',RSqlAfModeleTache.FindField ('AFM_LIBELLETACHE1').AsString);
     UneTob.AddChampSupValeur ('AFM_LIBELLETACHE2',RSqlAfModeleTache.FindField ('AFM_LIBELLETACHE2').AsString);
     UneTob.AddChampSupValeur ('AFM_FAMILLETACHE',RSqlAfModeleTache.FindField ('AFM_FAMILLETACHE').AsString);
     UneTob.AddChampSupValeur ('AFM_PLANNINGPERIOD',RSqlAfModeleTache.FindField ('AFM_PLANNINGPERIOD').AsString);
     UneTob.AddChampSupValeur ('AFM_MOISFIXE',RSqlAfModeleTache.FindField ('AFM_MOISFIXE').AsInteger);

     UneTob.AddChampSupValeur ('DPO_DATEREALISE',RSqlDpObligationRealise.FindField ('DPO_DATEREALISE').AsDateTime);
     UneTob.AddChampSupValeur ('DPO_UTILISATEUR',RSqlDpObligationRealise.FindField ('DPO_UTILISATEUR').AsString);
     UneTob.AddChampSupValeur ('DPO_ETAT',RSqlDpObligationRealise.FindField ('DPO_ETAT').AsInteger);
     UneTob.AddChampSupValeur ('DPO_COMMENTAIRE',RSqlDpObligationRealise.FindField ('DPO_COMMENTAIRE').AsString);
    end;
   Ferme (RsqlDpObligationRealise);
  end;
 Ferme (RSqlAfModeleTache);
 Result:=UneTob;
end;

//-----------------------------------------------------------------------------
//--- Nom   : DonnerLibelleTypeDeclaration
//--- Objet :
//-----------------------------------------------------------------------------
function DonnerLibelleTypeDeclaration (TypeDeclaration : String) : String;
begin
 if (TypeDeclaration='DEC') then Result:='DECLARATION';
 if (TypeDeclaration='PAI') then Result:='PAIEMENT';
 if (TypeDeclaration='OPT') then Result:='OPTION';
 if (TypeDeclaration='CON') then Result:='CONSEIL';
 if (TypeDeclaration='ASS') then Result:='ASSEMBLEE';
end;

//-----------------------------------------------------------------------------
//--- Nom   : DonnerFichierTelecharger
//--- Objet :
//-----------------------------------------------------------------------------
function DonnerFichierTelecharger (UneTob : Tob) : Integer;
var Position, Indice : Integer;
    DateCourante     : TDateTime;
begin
 Position:=0;
 DateCourante:=IDate1900;

 for Indice:=0 to UneTob.Detail.count-1 do
  if (UneTob.Detail[Indice].GetValue ('DATEDISPO')>DateCourante) then
   begin
    Position:=Indice;
    DateCourante:=UneTob.Detail[Indice].GetValue ('DATEDISPO');
   end;

 Result:=position;
end;

//*****************************************************************************

//-----------------------------------------------------------------------------
//--- Nom   : ChargerParametreObligation
//--- Objet :
//-----------------------------------------------------------------------------
function ChargerParametreObligation (Mode : String) : Tob;
var UneTob       : Tob;
    SRefClient   : String;
begin
 UneTob:=TOB.Create('',Nil,-1);
 UneTob.AddChampSupValeur ('MODE', Mode);
 UneTob.AddChampSupValeur ('ACTIVE', GetParamSocSecur ('SO_OBGACTIVE', False));
 UneTob.AddChampSupValeur ('DATEGENERATION', GetParamSocSecur ('SO_OBGDATEGEN', iDate1900));
 UneTob.AddChampSupValeur ('NUMVERSION', GetParamSocSecur ('SO_OBGVERSION', '0000'));
 UneTob.AddChampSupValeur ('PERIODICITE', GetParamSocSecur ('SO_OBGPERIODICITE', 'MEN'));
 UneTob.AddChampSupValeur ('ALERTE', GetParamSocSecur ('SO_OBGPASALERTE', False));
 UneTob.AddChampSupValeur ('TELECHARGEDIRECT', GetParamSocSecur ('SO_OBGTELECHARGEDIRECT', False));
 UneTob.AddChampSupValeur ('REPERTOIREDEST',DonnerRepertoireTemp);
 // Les clients ont l'habitude de mettre leur code client dans l'interface
 // de téléassistance cegidupdate => c'est bien ce code client qu'on veut passer
 // pour récupérer le fichier des obligations (et non pas $PGI$, même si le fichier
 // est publié sous le code générique $PGI$)
 SRefClient := ReadRegistryCegidUpdateSt ('CLIENT');
 if SRefClient='$PGI$' then SRefClient := '';
 UneTob.AddChampSupValeur ('REFERENCECLIENT', SRefClient);
 UneTob.AddChampSupValeur ('REPERTOIREPRISECOMPTE',DonnerRepertoireTemp);
 Result:=UneTob;
end;

//-----------------------------------------------------------------------------
//--- Nom   : TelechargerListeObligation
//--- Objet :
//-----------------------------------------------------------------------------
function TelechargerListeObligation (RepertoireDest : String; NomFichier : String) :Boolean;
var Erreur : String;
    UneTob : Tob;
    Indice : Integer;
begin
 Result:=False;
 UneTob:=GetFileListFromCegidUpdate (NomFichier,'','BUR',Now,Erreur);

 if (Erreur='') then
  begin
   if (UneTob<>Nil) then
    begin
     if (UneTob.Detail.Count>0) then
      begin
       Indice:=DonnerFichierTelecharger (UneTob);
       UneTob.Detail [Indice].PutValue ('REPDEST',RepertoireDest);
       TFDownloadUpdate.TelechargerDirectFile(UneTob.Detail[Indice]);
       Result:=True;
      end
     else
      PgiInfo ('Aucun fichier en ligne à télécharger.');
    end
  end
 else
  PgiInfo (Erreur);
end;

//-----------------------------------------------------------------------------
//--- Nom   : ImporterListeObligation
//--- Objet :
//-----------------------------------------------------------------------------
function ImporterListeObligation (CheminFichierObligation : String; NomFichier : String; Mode : String) : Boolean;
var TobListeObligation : Tob;
    NumVersion : String;
begin
 Result:=False;
 SourisSablier;
 TobListeObligation:=ChargerListeObligation (CheminFichierObligation+NomFichier);
 if (TobListeObligation<>nil) then
  begin
   NumVersion:=GetParamSocSecur ('SO_OBGVERSION', '0000');
   if NumVersion='' then NumVersion := '0000';
   if ( TobListeObligation.GetValue ('VERSION') <= NumVersion ) then
    begin
     PgiInfo ('Pas de nouvelle version des obligations.');
//     SetParamSoc ('SO_OBGVERSION',TobListeObligation.GetValue ('VERSION'));
    end
   else
    begin
     InitialiserTableModeleTache (TobListeObligation);
     CalculerListeObligation (Now,PlusMois (FinDeMois (Now),17));
     SetParamSoc ('SO_OBGVERSION',TobListeObligation.GetValue ('VERSION'));
     SetParamSoc ('SO_OBGDATEGEN',StrToDate (TobListeObligation.GetValue ('JOUR')+'/'+TobListeObligation.GetValue ('MOIS')+'/'+TobListeObligation.GetValue ('ANNEE')));
     if (Mode='AUTO') then SetParamSoc ('SO_OBGPASALERTE',False);
     TobListeObligation.Free;
     Result:=True;
    end;
  end;
 SourisNormale;
end;

//-----------------------------------------------------------------------------
//--- Nom   : ChargerListeObligation
//--- Objet :
//-----------------------------------------------------------------------------
function ChargerListeObligation (CheminFichier : String) : Tob;
var UnFichier            : TFichierTexte;
    ChInformationFichier : String;
    ChChampTob,ChSource  : String;
    UneTob,UneTobEnreg   : Tob;
begin
 UneTob:=Nil;

 if Not CAT_OuvrirFichierTexte (UnFichier,CheminFichier,MODE_LIRE) then
  PgiInfo ('Impossible d''ouvrir le fichier '+CheminFichier)
 else
  begin
   //--- Récupère le nom des champs de la TOB
   CAT_LireFichierTexte (UnFichier,ChChampTob);
   ChInformationFichier:=StringReplace(ChChampTob,'''','"',[rfReplaceAll]);

   CAT_LireFichierTexte (UnFichier,ChChampTob);
   ChChampTob:=StringReplace(ChChampTob,'''','"',[rfReplaceAll]);

   UneTob:=TOB.Create('',Nil,0);
   UneTob.AddChampSupValeur ('VERSION',copy (ChInformationFichier,1,4));
   UneTob.AddChampSupValeur ('PERIODICITE',copy (ChInformationFichier,5,1));
   UneTob.AddChampSupValeur ('ANNEE',copy (ChInformationFichier,6,4));
   UneTob.AddChampSupValeur ('MOIS',copy (ChInformationFichier,10,2));
   UneTob.AddChampSupValeur ('JOUR',copy (ChInformationFichier,12,2));
   while not CAT_FinFichierTexte (UnFichier) do
    begin
     CAT_LireFichierTexte (UnFichier,ChSource);
     ChSource:=StringReplace(ChSource,'''','"',[rfReplaceAll]);

     UneTobEnreg:=TOB.Create('',UneTob,-1);
     UneTobEnreg.LoadFromSt(ChChampTob,Chr(9),ChSource,chr(9));
    end;

   CAT_FermerFichierTexte (UnFichier);
  end;

 Result:=UneTob;
end;

//-----------------------------------------------------------------------------
//--- Nom   : InitialiserTableModeleTache
//--- Objet :
//-----------------------------------------------------------------------------
procedure InitialiserTableModeleTache (UneTob : Tob);
var TobAfModeleTache : Tob;
    Indice           : Integer;
    ChCode,ChDate    : String;
    CreationOk       : Boolean;
begin
 //--- On efface les enregistrements de la table AFMODELETACHE concernant le domaine FISCAL/SOCIAL/JURIDIQUE
 ExecuteSql ('DELETE FROM AFMODELETACHE WHERE AFM_FAMILLETACHE="FIS" OR AFM_FAMILLETACHE="SOC" OR AFM_FAMILLETACHE="JUR"');

 for Indice:=0 to UneTob.detail.count-1 do
  begin
   TobAfModeleTache:=Tob.Create ('AFMODELETACHE',nil,-1);
   TobAfModeleTache.SelectDB ('',nil,True);
   TobAfModeleTache.InitValeurs;
   CreationOk:=True;

   if (Unetob.Detail [Indice].FieldExists ('CODE')) then TobAfModeleTache.PutValue ('AFM_MODELETACHE',Unetob.Detail [Indice].GetValue ('CODE'));
   if (Unetob.Detail [Indice].FieldExists ('LIBELLE')) then TobAfModeleTache.PutValue ('AFM_LIBELLETACHE1',copy (Unetob.Detail [Indice].GetValue ('LIBELLE'),1,70));
   if (Unetob.Detail [Indice].FieldExists ('LIBELLELONG')) then TobAfModeleTache.PutValue ('AFM_LIBELLETACHE2',copy (Unetob.Detail [Indice].GetValue ('LIBELLELONG'),1,70));
   if (Unetob.Detail [Indice].FieldExists ('DOMAINE')) then TobAfModeleTache.PutValue ('AFM_FAMILLETACHE',UpperCase (copy (Unetob.Detail [Indice].GetValue ('DOMAINE'),1,3)));
   if (Unetob.Detail [Indice].FieldExists ('TYPE')) then TobAfModeleTache.PutValue ('AFM_CHARLIBRE3',Unetob.Detail [Indice].GetValue ('TYPE'));
   if (Unetob.Detail [Indice].FieldExists ('COMMENTAIRE')) then TobAfModeleTache.PutValue ('AFM_COMMENTAIRE',Unetob.Detail [Indice].GetValue ('COMMENTAIRE'));
   if (Unetob.Detail [Indice].FieldExists ('CONDITION')) then TobAfModeleTache.PutValue ('AFM_DESCRIPTIF',Unetob.Detail [Indice].GetValue ('CONDITION'));

   if IsNumeric (Unetob.Detail [Indice].GetValue ('JOUR')) then TobAfModeleTache.PutValue ('AFM_MOISJOURFIXE',Unetob.Detail [Indice].GetValue ('JOUR'));
   if IsNumeric (Unetob.Detail [Indice].GetValue ('MOIS')) then TobAfModeleTache.PutValue ('AFM_MOISFIXE',Unetob.Detail [Indice].GetValue ('MOIS'));
   if IsValidDate (Unetob.Detail [Indice].GetValue ('DATEDEBUTVALIDITE')) then TobAfModeleTache.PutValue ('AFM_DATEDEBPERIOD',StrToDate (Unetob.Detail [Indice].GetValue ('DATEDEBUTVALIDITE'))) else TobAfModeleTache.PutValue ('AFM_DATEFINPERIOD',StrToDate ('01/01/1900'));
   if IsValidDate (Unetob.Detail [Indice].GetValue ('DATEFINVALIDITE')) then TobAfModeleTache.PutValue ('AFM_DATEFINPERIOD',StrToDate (Unetob.Detail [Indice].GetValue ('DATEFINVALIDITE'))) else TobAfModeleTache.PutValue ('AFM_DATEFINPERIOD',StrToDate ('31/12/2099'));

   if (Unetob.Detail [Indice].FieldExists ('PERIODICITE')) then TobAfModeleTache.PutValue ('AFM_PLANNINGPERIOD',UpperCase (copy (Unetob.Detail [Indice].GetValue ('PERIODICITE'),1,1)));
   if (TobAfModeleTache.GetValue ('AFM_PLANNINGPERIOD')='S') then
    begin
     TobAfModeleTache.PutValue ('AFM_PLANNINGPERIOD','M');
     TobAfModeleTache.PutValue ('AFM_MOISFIXE',StrToInt ('6'));
    end;
   if (TobAfModeleTache.GetValue ('AFM_PLANNINGPERIOD')='T') then
    begin
     TobAfModeleTache.PutValue ('AFM_PLANNINGPERIOD','M');
     TobAfModeleTache.PutValue ('AFM_MOISFIXE',StrToInt('3'));
    end;

   //--- Si la date annuelle d'obligation n'est correcte on n'enregitre pas l'obligation
   if (TobAfModeleTache.GetValue ('AFM_PLANNINGPERIOD')='A') then
    begin
     ChCode:=Unetob.Detail [Indice].GetValue ('CODE');
     ChDate:=Unetob.Detail [Indice].GetValue ('JOUR')+'/'+Unetob.Detail [Indice].GetValue ('MOIS')+'/2005';
     if (not IsValidDate (ChDate)) then
      begin
       CreationOk:=False;
       PgiInfo ('Impossible d''enregistrer l''obligation '+ChCode+', présence d''une date "'+ChDate+'" non valide.');
      end
     else
      TobAfModeleTache.PutValue ('AFM_DATEANNUELLE',StrToDate (Unetob.Detail [Indice].GetValue ('JOUR')+'/'+Unetob.Detail [Indice].GetValue ('MOIS')+'/2005'));
    end;

   TobAfModeleTache.PutValue ('AFM_UNITETEMPS','J');
   TobAfModeleTache.PutValue ('AFM_MODEGENE',StrToInt ('1'));
   TobAfModeleTache.PutValue ('AFM_ANNEENB',StrToInt ('1'));
   TobAfModeleTache.PutValue ('AFM_MOISMETHODE',StrToInt ('1'));

   if (CreationOk) then TobAfModeleTache.InsertDb (Nil);
   TobAfModeleTache.Free;
  end;
end;

//-----------------------------------------------------------------------------
//--- Nom   : CalculerListeObligation
//--- Objet :
//-----------------------------------------------------------------------------
procedure CalculerListeObligation (DateDebutPeriode : TDateTime;DateFinPeriode : TDateTime);
var TobListeObligation    : Tob;
    UneTob,UneTobEnreg    : Tob;
    Indice                : Integer;
    CreationOk            : Boolean;
    ChJour,ChMois,ChAnnee : String;
    ChDate                : String;
begin
 //--- On efface les enregistrements de la table DPOBLIGATIONTMP
 ExecuteSql ('DELETE FROM DPOBLIGATIONTMP WHERE DPT_DATEOBLIGATION>"'+UsDateTime (DateDebutPeriode)+'"');

 GenererObligation (PlusDate (DateDebutPeriode,1,'J'),DateFinPeriode,TobListeObligation);
 if (TobListeObligation<>nil) then
  begin
   UneTob:=Tob.Create ('',nil,-1);
   for Indice:=0 to TobListeObligation.Detail.Count-1 do
    begin
     CreationOk:=True;
     ChDate:=TobListeObligation.Detail [Indice].GetValue ('DATE');

     //--- Si la Date n'est correcte on n'enregistre pas l'obligation
     //--- ou on la corrige pour 29/02 ou 30/02 => remplacé par 28/02
     if (not IsValidDate (ChDate)) then
      begin
       ChJour:=Cat_ExtraireChaine (ChDate,'','/');
       ChMois:=Cat_ExtraireChaine (ChDate,'/','/');
       ChAnnee:=Cat_ExtraireChaine (ChDate,'/','');

       if (ChMois='02') and ((ChJour='29') or (ChJour='30')) then
        ChDate:='28/'+ChMois+'/'+ChAnnee
       else
        begin
         CreationOk:=False;
         PgiInfo ('Présence d''une date "'+ChDate+'" non valide.');
        end;
      end
     //--- Date valide
     else
      begin
      // MD 08/06/06 - On ne prend pas en compte les obligations générées en dehors
      // de la période demandée (notamment : cas des obligations trimestrielles :
      //    avec afm_planningPeriod=M afm_moismethode=1 afm_moisfixe=3 :
      // => elles se répartissent de manière calendaire par année,
      //    sans tenir compte de la date de début de génération    )
      if StrToDate(ChDate)<=DateDebutPeriode then
       CreationOk:=False;
      end;

     if (CreationOk) then
      begin
       UneTobEnreg:=Tob.Create ('DPOBLIGATIONTMP',UneTob,-1);
       UneTobEnreg.PutValue ('DPT_DATEOBLIGATION',StrToDate (TobListeObligation.Detail [Indice].GetValue ('DATE')));
       UneTobEnreg.PutValue ('DPT_CODEOBLIGATION',TobListeObligation.Detail [Indice].GetValue ('AFM_MODELETACHE'));
      end;
    end;

   UneTob.InsertDB (nil);
   UneTob.Free;
   TobListeObligation.Free;
  end;
end;

//-----------------------------------------------------------------------------
//--- Nom   : GenererObligation
//--- Objet :
//-----------------------------------------------------------------------------
procedure GenererObligation (Datedeb, datefin : Tdatetime;var  pTobMere :tob);
var
  i, jj           : Integer;
  vQr             : TQuery;
  vTob,vTobMere   : Tob;
  vSt             : String;
  vAFRegles       : TAFRegles;
  vListeJours     : TStringList;
  vRdDuree        : Double; // durée des interventions
  vInNbJoursDecal : Integer; // nombre de jours de decalage maximum
  vStMethodeDecal : String;   // méthode de décalage des jours
  vStUniteTemps   : String;
begin
  vRdDuree  := 1;
  vInNbJoursDecal := 0;

  // $$$ JP 03/11/05 - inutilisé, normal?
  //vAFRegles := TAFRegles.create;

  pTobMere := TOB.Create('obligation mere', nil, -1);

  //on charge les taches modeles avec les champs à mettre dans la TOB
  vSt :='SELECT AFM_MODELETACHE, AFM_LIBELLETACHE1,AFM_LIBELLETACHE2,AFM_CHARLIBRE3,';
  vSt := vSt + ' AFM_DESCRIPTIF,AFM_FAMILLETACHE FROM AFMODELETACHE WHERE AFM_MODELETACHE like "Z%"';
  vTob := tob.create('liste obli', nil,-1);
  try
   vQr := openSql(vSt, true);
   if not vQr.eof then
   begin
    vTob.LoadDetailDB('AFMODELETACHE','','',vQR,False);
      // pour chaque tache  modele
    for i := 0 to vTob.detail.count - 1 do
    begin
      // chargement des regles de la tache
      vAFRegles := TAFRegles.create;
     try
     vAFRegles.LoadDBReglesModele(Vtob.detail[i].getstring('AFM_MODELETACHE'),CinRemplace,DateDeb,DateFin);
        //calcul de la liste des jours

     VStMethodeDecal:='A';
     vListeJours := vAFRegles.GenereListeJours(vRdDuree, vInNbJoursDecal,VStMethodeDecal, VStUniteTemps,CInNoCompteur);
     if vListeJours <> nil then
        begin
        for jj := 0 to vListeJours.count - 1 do
          begin //alimentation TOB pour chaque jour calculé
           MoveCurProgressForm('');
           vtobmere:= tob.create ('',ptobmere,-1);
           vtobmere.AddChampSupValeur('AFM_MODELETACHE', Vtob.detail[i].Getstring ('AFM_MODELETACHE'));
           vtobmere.AddChampSupValeur('DATE',vlistejours[jj]);
          end;
        end;
     vListeJours.Free;
     finally
     vAFRegles.Free;
     end;
    end; //fin boucle sur tob des modèles
  end; // fin requête renseignée
  finally
    vtob.Free;
    ferme(vQr);
    FiniMoveProgressForm;
    end;
end;

end.

