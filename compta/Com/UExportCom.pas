{***********UNITE*************************************************
Auteur  ...... : ENTRESSANGLE Manon
Créé le ...... : 20/04/2005
Modifié le ... :   /  /
Description .. : Module de traitement  en export, Compatible EAGL
Mots clefs ... :
*****************************************************************}

unit UExportCom;

interface

uses
windows,classes,
{$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
forms, UtilTrans, UTOB, RecordCom, HEnt1, HCTrls, stdctrls, Sysutils,
Paramsoc, RappType, ed_tools, uBOB,
{$IFNDEF EAGLSERVER}
 hmsgbox,
{$ELSE}
 HDebug, esession, ULibCpContexte,uWA, galMainSauveDossier,
 UGedFiles,
{$ENDIF}
Ent1, ULibEcriture, UTOZ, uYFILESTD, FileCtrl,
{$IFDEF SCANGED}
galSystem,
{$ENDIF}
ULibBonapayer,
UtilGed;

type

  TInfoEvent = function(Chaine: string; Listecom : TListBox) : Boolean of object;

  TExport = class
      private
      BaseCommune          : Boolean;
      ExclureEnreg,stArg   : string;
      Pdos,Aveclettrage    : Boolean;
      LISTEEXPORT          : TListBox;
      FichierImp           : string;
      PECRBUD,BANA,PJRL    : Boolean;
      EtablissUnique,BGed  : Boolean;
      ListeCpteGen         : TList;
      AUXILIARE,OkExport   : Boolean;
      // BVE 28.08.07 Suivi des validations
{$IFDEF CERTIFNF}
      PourArchivage        : Boolean;
{$ENDIF}
      GrandNbGene, GrandNbAux,PSECTIONana : Boolean;
      TOBTiers                            : TOB;
      HISTOSYNCHROValue,WhereEcriture     : string;
      Trans                               : TFTransfertcom;
      BNORMAL, BSIMULE, BSITUATION        : Boolean;
      BIFRS, BPREVISION, BREVISION        : Boolean;
      TYPEECRX,NATUREJRL,ECRINTEGRE       : string;
      BExportStat                         : TCheckBoxState;
      FAffInfo                            : TInfoEvent ;
      DateModif1,DateModif2               : string;
      AnnulValidation,Zippe               : Boolean;
      CPDaterecep                         : TDateTime;
{$IFDEF EAGLSERVER}
      FicIE                               : TextFile;
      FichierResultat,WhatAppli           : string;
      Nodossier                           : string;
      PathDos, PathStd, PathDat           : string;
{$ENDIF}
      CPTEDEBUT, CPTEFin, WhereAvance     : string;
      ListeFichierJoint,CorrepMdp         : string;
      GerBAP, PVentil,CorrepMdr           : Boolean;
      WhereNature                         : string;   // fiche 10431
      procedure InitTrans (TE : TOB);
      Procedure AffMessage(Mess : string);
      Procedure RemplirEnteteexport;
      procedure SauvegardeQueParametre(var F: TextFile);
      procedure ExportBudEcr(var F: TextFile);
      procedure RenWhereEcr( var Where : string; Ext : string);
      procedure ChargementTableLibre(ListeLibre : TList;Where : string);
      procedure ChargementTableSection(ListeSection : TList;Where : string);
      procedure Chargementregle(Listeregle : TList;Where : string);
      procedure ChargementCpte(ListeCpteGen : TList;Where : string; ForcerWhere : Boolean=FALSE);
      FUNCTION  ChargeRib(ListeRib : TList) : Boolean;
      procedure EcritureCpteauxGrandNombre(var F: TextFile);
      procedure ChargementCpteaux(ListeCpteaux : TList; Where : string; ForcerWhere : Boolean=FALSE);
      procedure EcrTiersCompl(var F: TextFile; Where : string);
      Procedure AlimJalB(LJB : TList; Where : string) ;
      procedure ChargementMouvement(var ListeMouvt : TList;Where, FormatEnvoie : string;Var F : TextFile; pana : string; ListeJalB : TList);
      procedure ChargementBalance(ListeMouvt : TList; Var F : TextFile);
      FUNCTION CHERCHEUNMODE (Mode : String3) : Char ;                                                                                               
      procedure ChargementMouvementAnalytique(var ListeMouvt : TList;Where,FormatEnvoie : string; Tdev : TList;Var F : TextFile; Var Indice : integer; LT9: string ='');
      procedure RemplirEcritureBal (var TMvt : PListeMouvt; Compte, comptecol, nature, libelle : string; var MttSolde : double);
      procedure Traitement_Ventil (var TMvt : PListeMouvt; var ListeMouvt : TList; Exo,Compte, Collectif, nature, libelle, Jrl : string);
      function  CreatWhereParam (Ext : string) : string;
      procedure ChargementParamlibre(Where : string; Var TobParam : TOB);
      Procedure ExportLesIMMOS;
      Procedure ExportLaCompta;
      procedure AddToBob(TStd: TOB; stSQL : string; sDomaine: string='C');
{$IFDEF EAGLSERVER}
      procedure FileClickzip(FFile : string='');
      Function  AfficheComExport(Chaine: string; Listecom : TListBox) : Boolean;
      procedure DownloadFileExport(var TD : TOB);
      function  ExportBobEagl (Fich : string; var TB : TOB) : Boolean;
      procedure RenseigneLesPath;
{$ENDIF}
      public

      Property  OnAfficheListeCom  : TInfoEvent read FAffInfo  write FAffInfo ;
      Procedure ExportComSx (TobExport : TOB; SArg : string);

  end;

Function LanceExport (var RequestTOB : TOB; Arg : string; AffInfo    : TInfoEvent; var PExport : TExport ; Archivage : Boolean = false) : Boolean;
{$IFDEF EAGLSERVER}
Function LanceZippeFile (var RequestTOB : TOB; Arg : string; AffInfo : TInfoEvent; var PExport : TExport) : Boolean;
{$ENDIF}
implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  CPTypeCons,
  {$ENDIF MODENT1}
  Constantes;

Function LanceExport (var RequestTOB : TOB; Arg : string; AffInfo : TInfoEvent; var PExport : TExport ; Archivage : Boolean = false) : Boolean;
{$IFDEF EAGLSERVER}
var
sTmp          : string;
Listefichiers : TStringList;
i             : integer;
{$ENDIF}
begin
  With PExport do
  begin
    if PExport = nil then PExport := TExport.Create;
    // BVE 28.08.07 Suivi des validations
{$IFDEF CERTIFNF}
    PourArchivage := Archivage;
{$ENDIF}
    OnAfficheListeCom     := Affinfo;
    ExportComSx(RequestTOB, Arg);
{$IFDEF EAGLSERVER}
    if (not Zippe) and (not BGed) then
           DownloadFileExport(RequestTOB)
    else
    begin
        if Zippe then FileClickzip;
        if ListeFichierJoint <> '' then
        begin
              sTmp     := ListeFichierJoint;
              Listefichiers := TStringList.Create;
              while ( length(sTmp) <> 0 ) do Listefichiers.add(ReadTokenPipe(sTmp,';'));
              for i:=0 to Listefichiers.Count-1 do
                  if (Listefichiers.Strings[i]<> '') then
                     FileClickzip(Listefichiers.Strings[i]);
              Listefichiers.free;
        end;
    end;
    if BGed then
    begin
            FileClickzip(ExtractFileDir(Trans.FichierSortie));
            RemoveInDir2(GetEnvVar('TEMP'), FALSE, FALSE, FALSE,'\Agl*.pdf');
    end;
{$ENDIF}
    free;
  end;
  Result := TRUE;
end;

{$IFDEF EAGLSERVER}

procedure TExport.RenseigneLesPath;
var
Nodisque             : integer;
Q                    : TQuery;
begin
    Nodisque := 0;
    Q := OpenSQL ('SELECT DOS_NODISQUE FROM DOSSIER WHERE DOS_NODOSSIER="' + nodossier + '"', TRUE);
    if not Q.EOF then
       Nodisque := Q.FindField('DOS_NODISQUE').asinteger;
    Ferme (Q);
    if (Nodisque >=1) and  (Nodisque<=8) then
       PathDos := TCPContexte.GetCurrent.MySession.Partages[Nodisque] + '\D' + nodossier
    else
       PathDos                  := TCPContexte.GetCurrent.MySession.dospath +'\D' + nodossier;

    PathStd                  := TCPContexte.GetCurrent.MySession.STDPath;
    PathDat                  := TCPContexte.GetCurrent.MySession.DatPath;
end;

// SI utilisation en base
Function LanceZippeFile (var RequestTOB : TOB; Arg : string; AffInfo : TInfoEvent; var PExport : TExport) : Boolean;
var
CodeRetour           : integer;
Filearchive          : string;
RepertoireRoot       : string;
begin
 Result := TRUE;
  With PExport do
  begin
    if PExport = nil then PExport := TExport.Create;
    Trans.FichierSortie       := RequestTOB.Detail.Items[RequestTOB.detail.count-1].getValue('FichierSortie');
    Zippe                     := RequestTOB.Detail.Items[RequestTOB.detail.count-1].GetValue('ZIPPE');
    WhatAppli                 := RequestTOB.GetValue('APPLICATION');

    nodossier := RequestTOB.GetValue('DOSSIER');
    RenseigneLesPath;

    if Zippe then
    begin
     Filearchive := ExtractFileName(Trans.FichierSortie);
     RepertoireRoot := RenseigneWrootPath(WhatAppli)+'\'+Filearchive;

     CodeRetour     :=  AGL_YFILESTD_IMPORT (RepertoireRoot, '0011', Filearchive, '.ZIP', 'COM', 'EXP', 'TMP');
     if CodeRetour = -1 then
     begin
         Result := TRUE;
         DeleteFile (RepertoireRoot);
     end
     else
          Result := FALSE
    end;
    free;
  end;
end;

procedure TExport.DownloadFileExport(var TD : TOB);
var
  FS      : TFileStream;
  St      : string;
begin
        if TD <> nil then TD := TOB.Create('', nil, -1);
        if not FileExists(Trans.FichierSortie) then exit;
        FS := TFileStream.Create(Trans.FichierSortie, fmOpenRead or fmShareDenyWrite);
        SetLength(st, FS.Size);
        FS.Seek(0, 0);
        FS.Read(PChar(st)^, FS.size);
        TD.AddChampSupValeur('DATA', St);
        FS.Free;
        LanceRemoveDirectory (WhatAppli);
end;

{$ENDIF}



procedure TExport.InitTrans (TE : TOB);
var
L1                          : TOB;
{$IFDEF EAGLSERVER}
RepertoireRoot              : string;
{$ENDIF}
WT,WG                       : string;
begin
     BaseCommune         := TE.GetValue('BaseCommune');
     L1                  := TE.Detail.Items[TE.detail.count-1];
     Trans.balance       := L1.getValue('Balance');
     Trans.complement    := L1.getValue('Complement') ;
     Trans.Dateecr1      := L1.getValue('DateEcr1') ;
     Trans.Dateecr2      := L1.getValue('DateEcr2') ;
     Trans.Etabi         := L1.getValue('Etabli') ;
     Trans.Exo           := L1.getValue('Exo') ;
     Trans.Exporte       := L1.getValue('Exporte') ;
     Trans.FichierSortie := L1.getValue('FichierSortie') ;
     Trans.Jr            := L1.getValue('Jr') ;
     Trans.naturejal     := L1.getValue('naturejal') ;
     Trans.pana          := L1.getValue('pana') ;
     Trans.Pgene         := L1.getValue('Pgene') ;
     Trans.pjournaux     := L1.getValue('pjournaux') ;
     Trans.psection      := L1.getValue('psection') ;
     Trans.ptiers        := L1.getValue('ptiers') ;
     Trans.ptiersautre   := L1.getValue('ptiersautre') ;
     Trans.Serie         := L1.getValue('Serie') ;
     Trans.Synch         := L1.getValue('Synch') ;
     Trans.Typearchive   := L1.getValue('Typearchive') ;
     Trans.TypeFormat    := L1.getValue('TypeFormat') ;
     Trans.Tlibre        := L1.getValue('Tlibre') ;
     if (Trans.Typearchive = 'DOS') then Trans.Tlibre := TRUE;

     PDos                := L1.getValue('PDos') ;
     Aveclettrage        := L1.getValue('Aveclettrage') ;
     ExclureEnreg        := L1.getValue('ExclureEnreg');
     FichierImp          := L1.getValue('FichierImp');
     PECRBUD             := L1.getValue('PECRBUD');
     BANA                := L1.GetValue('BANA');
     EtablissUnique      := L1.GetValue('EtablissUnique');
     AUXILIARE           := L1.GetValue('AUXILIARE');
     PSECTIONana         := L1.GetValue('PSECTIONana');
     BNORMAL             := L1.GetValue('BNORMAL');
     BSIMULE             := L1.GetValue('BSIMULE');
     BSITUATION          := L1.GetValue('BSITUATION');
     BIFRS               := L1.GetValue('BIFRS');
     BPREVISION          := L1.GetValue('BPREVISION');
     BREVISION           := L1.GetValue('BREVISION');
     TYPEECRX            := L1.GetValue('TYPEECRX');
     BExportStat         := L1.GetValue('STATE');
     WhereEcriture       := '';
     DateModif1          := L1.GetValue('DATEMODIF1');
     DateModif2          := L1.GetValue('DATEMODIF2');
     HISTOSYNCHROValue   := L1.GetValue('HISTOSYNCHROValue');  // FQ 10291 - CA - 03/11/2005
     WhereAvance         := L1.Getvalue ('WHEREAVANCE');
     WhereNature         := L1.Getvalue ('WHERENATURE'); // fiche 10431

     RenWhereEcr(WhereEcriture,'E');
     if (Trans.Typearchive = 'JRL') or (Trans.Typearchive = 'SYN') then
     begin
       WT :=' where t_auxiliaire in (select distinct(e_auxiliaire) from ecriture where '+ WhereEcriture + ') ';
       WG :=' where g_general in (select distinct(e_general) from ecriture where '+ WhereEcriture + ') ';
     end;

     InitDossierEnNombre(GrandNbGene, GrandNbAux, WG, WT);
     HISTOSYNCHROValue   := L1.GetValue('HISTOSYNCHROValue');
     NATUREJRL           := L1.GetValue('NATUREJRL');
     ECRINTEGRE          := L1.GetValue('ECRINTEGRE');
     PJRL                := L1.GetValue('PJRL');
     AnnulValidation     := L1.GetValue('ANNULVALIDATION');
     CPDaterecep         := GetParamSocSecur ('SO_CPRDDATERECEPTION', '');
     Zippe               := L1.GetValue('ZIPPE');
     BGed                := L1.GetValue('BGed') ;
     PVentil             := L1.Getvalue('PVentil');

     ListeFichierJoint := L1.Getvalue('ListeFichierJoint');
{$IFDEF EAGLSERVER}
     Nodossier := TE.GetValue('DOSSIER');
     WhatAppli := TE.GetValue('APPLICATION');
     RenseigneLesPath;
     RepertoireRoot := RenseigneWrootPath(WhatAppli);
     if not DirectoryExists(RepertoireRoot) then LanceCreatDir(WhatAppli);

     Trans.FichierSortie := RepertoireRoot+ '\'+ ExtractFileName(L1.getValue('FichierSortie'));

     if ((L1.GetValue('ModePCL')='1')) then V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxPCL] ;
     if not assigned(OnAfficheListeCom) then    // appel par process serveur
     begin
       OnAfficheListeCom := AfficheComExport;
       FichierResultat   := RepertoireRoot + '\ListeCom'+FormatDateTime(Traduitdateformat('yyyymmddhhnn'),NowH)+'.txt';
       AssignFile(FicIE,  FichierResultat) ;
       Rewrite(FicIE) ;
     end;
{$ENDIF}
     CPTEDEBUT := L1.GetValue('CPTEDEBUT');
     CPTEFIN   := L1.GetValue('CPTEFIN');
     OkExport := TRUE;
     if (Trans.Serie = 'S5') then
        GerBAP := (ExisteTypeVisa (''))
     else
        GerBAP := FALSE;
     CorrepMdp := '';
     if not ExisteSQL ('Select CR_CORRESP from CORRESP where CR_TYPE="MDP"') then
        CorrepMdp := 'FALSE';
     CorrepMdr := ExisteSQL ('Select CR_CORRESP from CORRESP where CR_TYPE="MDR"');
end;

Procedure TExport.AffMessage(Mess : string);
begin
{$IFNDEF EAGLSERVER}
        PgiBox(Mess+'#10','Envoyer') ;
{$ELSE}
        //ddWriteLN(Mess);
        cWA.MessagesAuClient('COMSX.EXPORT','',Mess) ;
{$ENDIF}
end;

Procedure TExport.RemplirEnteteexport;
begin
          Trans.Synch := '';
          if PDos  then                      Trans.complement := 'X'
          else                                     Trans.complement := '-';

          if (Trans.Typearchive = 'DOSS') then Trans.Synch := 'SYN'
          else
          if (Trans.Typearchive = 'JRL') and Aveclettrage then
          begin
               Trans.complement := ''; Trans.Synch := 'REJ'
          end;

end;


Procedure TExport.ExportComSx (TobExport : TOB; SArg : string);
var
ListeJalB, ListeMouvt, ListeCpteaux : TList;
ListeLibre,ListeSection,ListeRib    : TList;
Listeregle              : TList;
FichierIE               : TextFile;
Where                   : string;
Jourx                   : string;
Stt                     : string;
WhereEcr                : string;
TOBL                    : TOB;
begin
    stArg := SArg;   

    InitTrans (TobExport);
    if Trans.Serie = 'SAGE' then
    begin
            AssignFile(FichierIE,Trans.FichierSortie) ;
            try Rewrite(FichierIE) ; except end;
            if IoResult<>0 then
              begin
                AffMessage('Impossible d''écrire dans le fichier #10'+Trans.FichierSortie) ;
                Exit ;
              end ;
            try
                  Stt :=GetParamSocSecur('SO_LIBELLE', '') ; Stt :=Format_String(Stt ,30) ;
                  Writeln(FichierIE, Stt) ;
                  Where := '';
               // pour les écritures
                  ListeMouvt:=TList.Create;
                  RenWhereEcr(Where,'E');
                  AfficheListeCom('Chargement des écritures', LISTEEXPORT);
                  ChargementMouvement(ListeMouvt,Where,Trans.TypeFormat,FichierIE, Trans.pana, nil);
                  AfficheListeCom('Export des écritures', LISTEEXPORT);

                  EcritureMouvement(ListeMouvt,FichierIE, Trans.Serie, CorrepMdp);
                  ListeMouvt.Free ;
            finally
                CloseFile(FichierIE);
                {$IFDEF EAGLSERVER}
                CloseFile(FicIE);
                {$ENDIF}
            end ;
            exit;
    end
    else
    begin
            if Trans.Serie = 'IMMO' then
            begin
             ExportLesIMMOS;
{$IFDEF EAGLSERVER}
           CloseFile(FicIE);
{$ENDIF}
             exit;
            end
            else
            begin
                  if Trans.Serie = 'COMPTA' then
                  begin
                       ExportLaCompta;
                       {$IFDEF EAGLSERVER}
                       CloseFile(FicIE);
                       {$ENDIF}
                       exit;
                  end;
            end;
    end;

    AssignFile(FichierIE, Trans.FichierSortie) ;
    try Rewrite(FichierIE) ; except end;
    if IoResult<>0 then
    begin
        AffMessage('Impossible d''écrire dans le fichier'+Trans.FichierSortie) ;
        Exit ;
    end ;
    try
        // entete
          RemplirEnteteexport;

          Ecritureentete(FichierIE, Trans.Typearchive, Trans.TypeFormat, Trans.Serie, Trans.complement, Trans.Synch);

        // pour les paramètres généraux
          if (ExclureEnreg = '') or (pos('PARAM', ExclureEnreg) = 0)  then
          begin
            OnAfficheListeCom('Paramètres généraux', LISTEEXPORT);
            Ecrituregeneraux(FichierIE);
          end;

        // pour les exercices
          if (ExclureEnreg = '') or (pos('EXO', ExclureEnreg) = 0) then
          begin
            OnAfficheListeCom('Exercice', LISTEEXPORT);
            Where := '';
            if (Trans.Exo <> '') then
               Where := ' Where ' +RendCommandeExo ('EX',Trans.Exo);
            EcritureExercice(FichierIE, Where);
          end;
          Where := '';

          // si l'export uniquement des paramètres dossiers
          if PDos  then
          begin
               SauvegardeQueParametre(FichierIE);
               exit;
          end;
        // Ecriture budgétraire
        if PECRBUD and (Trans.Typearchive = 'JRL') then
        begin
               ExportBudEcr(FichierIE);
               exit;
        end;

        if (Trans.Typearchive = 'JRL') or (Trans.Typearchive = 'SYN') then RenWhereEcr(WhereEcr,'E');

        if (Trans.Serie = 'S1') then
          EcritureChoixCod(FichierIE, Trans.Serie);

        if (Trans.Serie <> 'S1') then
        begin
          // pour table libre
          if (ExclureEnreg = '') or (pos('TL', ExclureEnreg) = 0) then
          begin
             OnAfficheListeCom('Tables Libres', LISTEEXPORT);
             // table personnalisée
             if Trans.Tlibre then
             begin
                  ListeLibre := TList.Create;
                  ChargementTableLibre(ListeLibre, ' left join CHOIXCOD on CC_CODE=NT_TYPECPTE where cc_abrege="X"');
                  Ecrituretablelibre(ListeLibre,FichierIE);
                  ListeLibre.Free ;
                  EcritureChoixCod(FichierIE);
                  ChargementParamlibre(' Where CC_ABREGE="X" ', TOBL);
                  EcritureParamlib(FichierIE, TOBL);
                  if TOBL <> nil then TOBL.free;
                  EcritureDesRessource(FichierIE, '');
             end;
           end;
             // section analytique
           if ((Trans.Pana = 'X') and (Trans.balance <> TRUE))  or (BANA and (Trans.balance = TRUE)) then
           begin
             if (ExclureEnreg = '') or (pos('SSA', ExclureEnreg) = 0) then
             begin
               OnAfficheListeCom('Sections analytiques', LISTEEXPORT);
               ListeSection := TList.Create;
               ChargementTableSection(ListeSection, '');
               Ecrituretablesection(ListeSection,FichierIE);
               ListeSection.Free ;
             end;
           end;
             // établissement
          if (ExclureEnreg = '') or (pos('ETB', ExclureEnreg) = 0) then
          begin
              OnAfficheListeCom('Etablissements', LISTEEXPORT);
              Where := '';
              if (Trans.Typearchive = 'JRL')
              or (Trans.Typearchive = 'SYN') then
                   Where :=' where ET_ETABLISSEMENT in (select distinct(E_ETABLISSEMENT) from ecriture where '+ WhereEcr + ') '
              else
                   Where := '';

              if (EtablissUnique) then
                 EcritureEtablissement(FichierIE, ' Where ET_ETABLISSEMENT="'+GetParamSocSecur ('SO_ETABLISDEFAUT', '')+'"')
              else
                 EcritureEtablissement(FichierIE, Where);
              Where := '';
          end;
          if (ExclureEnreg = '') or (pos('MDP', ExclureEnreg) = 0) then
            EcritureModepaiement(FichierIE, CorrepMdp);

         end;

       // règlement
         if (ExclureEnreg = '') or (pos('MDR', ExclureEnreg) = 0) then
         begin
             OnAfficheListeCom('Réglements', LISTEEXPORT);
             Listeregle := TList.Create;
             Chargementregle(Listeregle, '');
             Ecrituretableregle(Listeregle,FichierIE, Trans.Serie, stArg);
             Listeregle.Free ;
         end;

         if (ExclureEnreg = '') or (pos('DEV', ExclureEnreg) = 0) then
         begin
              OnAfficheListeCom('Devise', LISTEEXPORT);
              Ecrituredevise(FichierIE);
         end;
         
        // BVE 28.08.07 Suivi des validations
{$IFDEF CERTIFNF}
        if PourArchivage then
           EcritureSuiviValidation(FichierIE);
{$ENDIF}

        // régime TVA
        if (Trans.Serie <> 'S1')
         and ((ExclureEnreg = '') or (pos('REG', ExclureEnreg) = 0)) then
        begin
             OnAfficheListeCom('TVA', LISTEEXPORT);
             Ecrituretva(FichierIE);
        end;

        // section analytique
        if ((Trans.Pana = 'X') and (Trans.balance <> TRUE)) or (BANA  and (Trans.balance = TRUE)) then
        begin
         if ((ExclureEnreg = '') or (pos('SAT', ExclureEnreg) = 0)) then
         begin
              OnAfficheListeCom('Sections analytiques', LISTEEXPORT);
              Where := '';
              if(Trans.Typearchive = 'SYN') then // table histo
              begin
                if ExisteSQL ('SELECT CHT_NOMTABLE FROM CPHISTOPARAM where (CHT_CODE="") and (CHT_NOMTABLE="SECTION" )') then
                   Where := ' LEFT JOIN CPHISTOPARAM on (CHT_CODE="") where (CHT_NOMTABLE="SECTION" )'
                else
                  if ExisteSQL ('SELECT S_SECTION FROM SECTION  LEFT JOIN CPHISTOPARAM on (CHT_CODE=S_SECTION) where (CHT_NOMTABLE="SECTION" )') then
                    Where := ' LEFT JOIN CPHISTOPARAM on (CHT_CODE=S_SECTION) where (CHT_NOMTABLE="SECTION" )';
              end;
              EcritureSectionana(FichierIE, Where);
         end;
        end;

        // J_COMPTEURNORMAL
        if (Trans.Serie <> 'S1') and (Trans.balance <> TRUE) then
        begin
             // souche
             if (ExclureEnreg = '') or (pos('SOU', ExclureEnreg) = 0) then
             begin
                 OnAfficheListeCom('Souches', LISTEEXPORT);
                 Where := '';
                 if Trans.Jr <> ''  then
                 begin
                       Jourx := Trans.Jr;
                       Where := ' Where sh_souche in (select distinct(j_compteurnormal) '+
                                '  from journal ' +
                                ' Where ' + RendCommandeJournal('J',Jourx) + ')';
                 end
                 else Where := ' Where SH_TYPE="CPT" ';
                 EcritureSouche(FichierIE, Where);
                 Where := '';
             end;

             // Correspondance
             if (ExclureEnreg = '') or (pos('CORR', ExclureEnreg) = 0) then
             begin
                  OnAfficheListeCom('Correspondants', LISTEEXPORT);
                  EcritureCorresp(FichierIE);
                  EcritureCorrespimpot(FichierIE);
             end;
             // relance
             if (ExclureEnreg = '') or (pos('REL', ExclureEnreg) = 0) then
                EcritureRelance(FichierIE, '');
                        // fiche 10628
             if PVentil or (Trans.Typearchive = 'DOS') then // ventilation type
                EcritureVentilType(FichierIE);
        end;

        // pour les comptes généraux
        if (Trans.Pgene = 'X')  or (Trans.Typearchive = 'SYN') then
        begin
          if(Trans.Typearchive = 'SYN') then
              Where := ' Where G_DATECREATION >="'+ USDATETIME(CPDaterecep) +'"';
          OnAfficheListeCom('Comptes généraux', LISTEEXPORT);
          if ListeCpteGen = nil then
          begin
               ListeCpteGen:=TList.Create;
               ChargementCpte(ListeCpteGen,Where);
               if(Trans.Typearchive = 'SYN') then // table histo
               begin
                    if ExisteSQL ('SELECT CHT_NOMTABLE FROM CPHISTOPARAM where (CHT_CODE="") and (CHT_NOMTABLE="GENERAUX" )') then
                    begin
                       Where := ' LEFT JOIN CPHISTOPARAM on (CHT_CODE="") where (CHT_NOMTABLE="GENERAUX" )';
                       ChargementCpte(ListeCpteGen, Where, TRUE);
                    end else
                    if ExisteSQL ('SELECT G_GENERAL FROM GENERAUX  LEFT JOIN CPHISTOPARAM on (CHT_CODE=G_GENERAL) where (CHT_NOMTABLE="GENERAUX" )') then
                    begin
                       Where := ' LEFT JOIN CPHISTOPARAM on (CHT_CODE=G_GENERAL) where (CHT_NOMTABLE="GENERAUX" )';
                       ChargementCpte(ListeCpteGen,Where, TRUE);
                    end;
               end;
          end;
          // Fiche 22010
          if (ExclureEnreg = '') or (pos('CGN', ExclureEnreg) = 0)  then
            EcritureCpptegen(ListeCpteGen,FichierIE);
          ListeCpteGen.Free ;
          Where := '';
        end;
        if (ExclureEnreg = '') or (pos('RIB', ExclureEnreg) = 0) then
        begin
            ListeRib:=TList.Create;
            ChargeRib(ListeRib);
            EcritureRib(ListeRib,FichierIE);
            ListeRib.free;
        end;
        if (Trans.Serie <> 'S1') then
        begin
             if (ExclureEnreg = '') or (pos('BQC', ExclureEnreg) = 0) then
                EcritureBanquecp(FichierIE);
             if (ExclureEnreg = '') or (pos('BQE', ExclureEnreg) = 0) then
             begin
                  EcritureBanques(FichierIE);
                  EcritureReleveBanque(FichierIE);
                  EcritureLigneReleveBanque(FichierIE);
             end;
        end
        else
        begin
             if (ExclureEnreg = '') or (pos('BQE', ExclureEnreg) = 0) then
             begin
                  EcritureReleveBanque(FichierIE);
                  EcritureLigneReleveBanque(FichierIE);
             end;
        end;
        if GerBAP then //BAP
        begin
             if (ExclureEnreg = '') or (pos('BAP', ExclureEnreg) = 0) then
             begin
                  EcritureBapTypeVisa(FichierIE);
                  EcritureBapCircuit(FichierIE);
             end;
        end;
        // pour les comptes tiers
        if ((Trans.ptiers = 'X') and (Trans.balance <> TRUE)) or
        (Trans.Typearchive = 'SYN') or
        (AUXILIARE and (Trans.balance = TRUE))then
        begin
          if(Trans.Typearchive = 'SYN') then
          Where := ' Where T_DATECREATION >="'+ USDATETIME(CPDaterecep) +'"';

          OnAfficheListeCom('Comptes Tiers', LISTEEXPORT);
          if GrandNbAux then //AJOUT ME 24-01-2005
                  EcritureCpteauxGrandNombre(FichierIE)
          else
          begin
                  ListeCpteaux:=TList.Create;
                  ChargementCpteaux(ListeCpteaux,Where);
                  if(Trans.Typearchive = 'SYN') then // table histo
                  begin
                       if ExisteSQL ('SELECT CHT_NOMTABLE FROM CPHISTOPARAM WHERE (CHT_CODE="") AND (CHT_NOMTABLE="TIERS" )') then
                       begin
                          Where := ' LEFT JOIN CPHISTOPARAM on (CHT_CODE="") where (CHT_NOMTABLE="TIERS" )';
                          ChargementCpteaux(ListeCpteaux,Where, TRUE);
                       end else
                       if ExisteSQL ('SELECT T_AUXILIAIRE FROM TIERS  LEFT JOIN CPHISTOPARAM on (CHT_CODE=T_AUXILIAIRE) where (CHT_NOMTABLE="TIERS" )') then
                       begin
                          Where := ' LEFT JOIN CPHISTOPARAM on (CHT_CODE=T_AUXILIAIRE) where (CHT_NOMTABLE="TIERS" )';
                          ChargementCpteaux(ListeCpteaux,Where, TRUE);
                       end;
                  end;
                  // fiche 22010
                  if (ExclureEnreg = '') or (pos('CAE', ExclureEnreg) = 0)  then
                     EcritureCppteaux(ListeCpteaux,FichierIE, stArg);
                  ListeCpteaux.Free ;
                  if (TOBTiers <> nil) and (stArg <> '') then TOBTiers.free;
          end;

                 // fiche 22010
          if (ExclureEnreg = '') or (pos('CAE', ExclureEnreg) = 0)  then
          begin
           Where := ' Where YTC_DATECREATION >="'+ USDATETIME(CPDaterecep) +'"';
           EcrTiersCompl(FichierIE, Where);
          end;
          Where := '';
        end;

       // pour les journaux
          if Trans.balance = TRUE then
          begin
               if (Trans.Jr = '') or (Trans.Jr = Traduirememoire('<<Tous>>'))  then
               begin
                    if (ctxPCL in V_PGI.PGIContexte) then
                       Trans.Jr := 'CO'
                    else
                       Trans.Jr := 'CO;AA';
               end;
          end;
{ ajout me 18-07-2005 E_PAQUETREVISION=1 uniquement dans le dossier pour envoyer en synchro
  deplacement de update avant la liste des journaux}
          If (Trans.Typearchive = 'SYN') Then
             UpdateComLettrage(0) ;

          ListeJalB:=TList.Create ;
          if Trans.Jr <> ''  then
          begin
               OnAfficheListeCom('Les journaux : ' + Trans.Jr, LISTEEXPORT);
               Jourx := Trans.Jr;
               Where := ' Where ' + RendCommandeJournal('J',Jourx);
          end
          else
          begin
              OnAfficheListeCom('Les journaux', LISTEEXPORT);
              if (Trans.Typearchive = 'JRL')
              or (Trans.Typearchive = 'SYN') then
              begin
                   Where :=' where (J_JOURNAL in (select distinct(E_JOURNAL) from ecriture where '+ WhereEcr + ')) ';
                   if (Trans.Serie <> 'S1') then // pour les ODA
                   Where := Where +
                   ' OR (J_JOURNAL in (select distinct(Y_JOURNAL) from analytiq where  Y_TYPEANALYTIQUE="X"))';
              end;
          end;
          AlimJalB(ListeJalB, Where) ;
          if(Trans.Typearchive = 'SYN') then // table histo
          begin
                if ExisteSQL ('SELECT CHT_NOMTABLE FROM CPHISTOPARAM where (CHT_CODE="") and (CHT_NOMTABLE="JOURNAUX" )') then
                begin
                   Where := ' LEFT JOIN CPHISTOPARAM on (CHT_CODE="") where (CHT_NOMTABLE="JOURNAUX" )';
                   AlimJalB(ListeJalB, Where);
                end else
                if ExisteSQL ('SELECT J_JOURNAL FROM JOURNAL  LEFT JOIN CPHISTOPARAM on (CHT_CODE=J_JOURNAL) where (CHT_NOMTABLE="JOURNAL" )') then
                begin
                   Where := ' LEFT JOIN CPHISTOPARAM on (CHT_CODE=J_JOURNAL) where (CHT_NOMTABLE="JOURNAL" )';
                   Where := Where + ' and (J_JOURNAL not in (select distinct(E_JOURNAL) from ecriture where '+ WhereEcr + ')) ';
                   AlimJalB(ListeJalB, Where);
                end;
          end;

          // Fiche 22010
          if (ExclureEnreg = '') or (pos('JAL', ExclureEnreg) = 0)  then
             EcritureJournal(ListeJalB, FichierIE, FALSE);
          Where := '';

       // pour les écritures
 {
 WHERE E_DATECOMPTABLE>='01/01/2001'
 AND E_DATECOMPTABLE<='12/31/2001'
 AND E_EXERCICE='001'
 AND E_ETABLISSEMENT='166'
 AND E_NUMEROPIECE>=0 AND E_NUMEROPIECE<=999999999
 AND E_JOURNAL>='ACH' AND E_JOURNAL<='VEN' AND E_QUALIFPIECE='N' AND E_EXPORTE='X'
 }

        // Sauvegarde des écritures
        if (Trans.balance <> TRUE) then
        begin
          ListeMouvt:=TList.Create;
          RenWhereEcr(Where,'E');
          OnAfficheListeCom('Chargement des écritures', LISTEEXPORT);
          if Trans.balance <> TRUE then
             ChargementMouvement(ListeMouvt, Where, Trans.TypeFormat, FichierIE, Trans.pana, ListeJalB);
          OnAfficheListeCom('Export des écritures', LISTEEXPORT);

          EcritureMouvement(ListeMouvt,FichierIE,Trans.TypeFormat,CorrepMdp);
          ListeMouvt.Free ;
        end;
        // Initialisation les informations historiques  // fiche 10181
        If (Trans.Typearchive = 'SYN') and ((ListeJalB <> nil) and (ListeJalB.Count > 0)) Then
        begin
                   if (HISTOSYNCHROValue = '0') then UpdateComLettrage(4)
                   else
                   if (HISTOSYNCHROValue = '1') then UpdateComLettrage(5)
                   else
                                                      UpdateComLettrage(3) ;
        end
        else
              if ((Trans.Serie = 'S1') and (Trans.Typearchive = 'DOS'))
              or ((Trans.Serie <> 'S1') and (Trans.Serie <> 'SISCO') and (Trans.Typearchive = 'DOSS'))then
                 UpdateComLettrage(6);

        If (Trans.Typearchive = 'SYN') or (Trans.Serie = 'S1') Then
           SetParamsoc ('SO_CPDERNEXOCLO','');
        if (Trans.balance = TRUE) and (not PDOS) then
        begin
               OnAfficheListeCom('Balance', LISTEEXPORT);
               ListeMouvt:=TList.Create;
               ChargementBalance(ListeMouvt, FichierIE);
               EcritureMouvement(ListeMouvt,FichierIE,Trans.TypeFormat,CorrepMdp);
               ListeMouvt.Free ;
        end;
        if ListeJalB <> nil then
        begin
             LiberJournal(ListeJalB);
             ListeJalB.Free ;
        end;
    finally
           CloseFile(FichierIE);
           if(Trans.Typearchive = 'SYN') then // table histo
             ExecuteSQL ('Delete from CPHISTOPARAM');

           OnAfficheListeCom('Export terminé ', LISTEEXPORT);
{$IFDEF EAGLSERVER}
           CloseFile(FicIE);
{$ENDIF}
    end ;
    if (Trans.Typearchive = 'SYN') then
       SetParamsoc ('SO_CPSYNCHROSX', TRUE);
    if (GetparamsocSecur ('SO_CPMODESYNCHRO', TRUE, FALSE) and (Trans.Typearchive = 'DOS'))  then
       SetParamsoc ('SO_FLAGSYNCHRO', 'SYN')
    else
       SetParamsoc ('SO_FLAGSYNCHRO', Trans.Typearchive);
    if (Trans.Serie = 'S1') and (Trans.Typearchive = 'DOS')
       and (GetParamSocSecur ('SO_CPDECLOTUREDETAIL', False)= TRUE) then
    GetParamSocSecur ('SO_CPDECLOTUREDETAIL', FALSE);

end;

function TExport.CreatWhereParam (Ext : string) : string;
var
datedeb,datefin : TDATETIME;
begin
         Result := '';
         datedeb := StrToDate(DATEMODIF1);
         datefin := StrToDate(DATEMODIF2);
         if datedeb = datefin then
            Result := ' Where ('+Ext+'_DATEMODIF >="'+ USTIME(StrToDate(DATEMODIF1)) +'"' +
            ' AND '+Ext+'_DATEMODIF < "'+ USTIME(datefin+1) +'")'
         else
             Result := ' Where ('+Ext+'_DATEMODIF >="'+ USTIME(StrToDate(DATEMODIF1)) +'"' +
             ' AND '+Ext+'_DATEMODIF <="'+ USTIME(StrToDate(DATEMODIF2)) +'")';
         if CPTEDEBUT <> '' then
         begin
               case Ext[1] of
                 'G' :
                  Result := Result +
                  ' And ('+Ext+'_GENERAL >="'+  CPTEDEBUT+'"' +
                  ' AND '+Ext+'_GENERAL <="'+ CPTEFIN +'")';
                 'T','Y' :
                  Result := Result +
                  ' And ('+Ext+'_AUXILIAIRE >="'+  CPTEDEBUT+'"' +
                  ' AND '+Ext+'_AUXILIAIRE <="'+ CPTEFIN +'")';
                  'S' :
                  Result := Result +
                  ' And ('+Ext+'_SECTION >="'+  CPTEDEBUT+'"' +
                  ' AND '+Ext+'_SECTION <="'+ CPTEFIN +'")';
                  'J' :
                  Result := Result +
                  ' And ('+Ext+'_JOURNAL >="'+  CPTEDEBUT+'"' +
                  ' AND '+Ext+'_JOURNAL <="'+ CPTEFIN +'")';
                 'B' : // compte budgétaires
                     case Ext[2] of
                          'G' :
                               Result := Result +
                               ' And ('+Ext+'_BUDGENE >="'+  CPTEDEBUT+'"' +
                               ' AND '+Ext+'_BUDGENE <="'+ CPTEFIN +'")';
                          'J' :
                               Result := Result +
                               ' And ('+Ext+'_BUDJAL >="'+  CPTEDEBUT+'"' +
                               ' AND '+Ext+'_BUDJAL <="'+ CPTEFIN +'")';
                          'S' : // section  BUDSECT
                               Result := Result +
                               ' And ('+Ext+'_BUDSECT >="'+  CPTEDEBUT+'"' +
                               ' AND '+Ext+'_BUDSECT <="'+ CPTEFIN +'")';
                     end;
               end;
         end;
   if (length(Ext) = 2) then
   begin
        if (Ext[2] <> 'J')  then
        if not(BExportStat=cbGrayed) then
             if BExportStat=cbChecked then Result := Result + ' AND '+Ext+'_EXPORTE="X"' else Result := Result +' AND '+Ext+'_EXPORTE<>"X"' ;
   end
   else
   begin
        if not(BExportStat=cbGrayed) and (Ext <> 'YTC') then
             if BExportStat=cbChecked then Result := Result + ' AND '+Ext+'_EXPORTE="X"' else Result := Result +' AND '+Ext+'_EXPORTE<>"X"' ;
   end;

   if (WhereNature <> '') then   // fiche 10431
   begin
    if (Ext[1] = 'J') or (Ext[1] = 'G') or (Ext[1] = 'T') then
    Result := Result + ' AND (' + WhereNature + ')';
   end;

end;

procedure TExport.SauvegardeQueParametre(var F: TextFile);
var
ListeJalB, ListeCpteaux, ListeSection : TList;
ListeLibre    : TList;
CWhere        : string;
TOBL          : TOB;
ListeRib      : TList;
begin
        if (Trans.Pgene = 'X') then
        begin
                if (not PECRBUD) then
                begin
                  OnAfficheListeCom('Comptes généraux', LISTEEXPORT);
                  CWhere :=  CreatWhereParam('G');
                  if ListeCpteGen = nil then
                  begin
                       ListeCpteGen:=TList.Create;
                       ChargementCpte(ListeCpteGen, CWhere);
                  end;
                  EcritureCpptegen(ListeCpteGen,F);
                  ExecuteSql('UPDATE GENERAUX SET G_EXPORTE="X" ' + CWhere) ;
                  ListeCpteGen.Free ;
                end
                else  // comptes bugetaires
                begin
                  OnAfficheListeCom('Comptes Budgétaires', LISTEEXPORT);
                  CWhere := CreatWhereParam('BG');
                  EcrireCpteBudgetaire(F, CWhere);
                  ExecuteSql('UPDATE BUDGENE SET BG_EXPORTE="X" ' + CWhere) ;

                end;
        end;

        if (Trans.ptiers = 'X') and (not PECRBUD) then
        begin
          OnAfficheListeCom('Comptes Tiers', LISTEEXPORT);
          ListeCpteaux:=TList.Create;
          CWhere := CreatWhereParam('T');
          ChargementCpteaux(ListeCpteaux, CWhere);
          EcritureCppteaux(ListeCpteaux,F, stArg);
          ExecuteSql('UPDATE TIERS SET T_EXPORTE="X" ' + CWhere) ;
          ListeCpteaux.Free ;
          // sauvegarde des tiers complémentaires
          EcrTiersCompl(F, CreatWhereParam('YTC'));
          if TOBTiers <> nil then TOBTiers.free;
          // ajout me 25/08/2005 export balance des ribs liés aux tiers
          ListeRib:=TList.Create;
          ChargeRib(ListeRib);
          EcritureRib(ListeRib, F);
          ListeRib.free;
        end;

       if (PSECTIONana) then
       begin
                if (not PECRBUD) then
                begin
                     OnAfficheListeCom('Sections analytiques', LISTEEXPORT);
                     ListeSection := TList.Create;
                     ChargementTableSection(ListeSection, '');
                     Ecrituretablesection(ListeSection,F);
                     EcritureSectionana(F, CreatWhereParam('S'));
                     ListeSection.Free ;

                end
                else
                begin
                   OnAfficheListeCom('Section Budgétaires', LISTEEXPORT);
                   CWhere := CreatWhereParam('BS');
                   EcritureSectionBudgetaire (F, CWhere);
                   ExecuteSql('UPDATE BUDSECT SET BS_EXPORTE="X" ' + CWhere) ;
                end;
       end;

       if (PJRL) then
       begin
                if (not PECRBUD) then
                begin
                      ListeJalB := TList.Create;
                      AlimJalB(ListeJalB, CreatWhereParam('J')) ;
                      EcritureJournal(ListeJalB, F);
                      ListeJalB.Free ;
                end
                else
                begin
                   OnAfficheListeCom('Journaux Budgétaires', LISTEEXPORT);
                   EcrireJalBudgetaire(F, CreatWhereParam('BJ'));
                end;
       end;
       // table personnalisée
       if Trans.Tlibre then
       begin
             ListeLibre := TList.Create;
             ChargementTableLibre(ListeLibre, '');
             Ecrituretablelibre(ListeLibre,F);
             ListeLibre.Free ;
             EcritureChoixCod(F);
             ChargementParamlibre(' Where CC_ABREGE="X" ', TOBL);
             EcritureParamlib(F, TOBL);
             if TOBL <> nil then TOBL.free;
             EcritureDesRessource(F, '');
       end;
       if PVentil then // ventilation type
             EcritureVentilType(F);

end;

procedure TExport.ExportBudEcr(var F: TextFile);
var
lStSQL    : string;
Q1        : Tquery;
Sens,WE   : string;
Ligne     : string;
Orderby   : string;
begin

     lStSQL:=CGetSQLFromTable('BUDECR', ['BE_BLOCNOTE','BE_EXERCICE']) ;
     RenWhereEcr(WE,'BE');
     lStSQL := lStSQL + ' Where '+ WE;
     Orderby := ' order by BE_BUDJAL,BE_EXERCICE,BE_DATECOMPTABLE,BE_BUDSECT,BE_AXE,BE_NUMEROPIECE,BE_QUALIFPIECE';

     Q1:=OpenSQL(lStSQL+ OrderBy, TRUE);
     While not Q1.EOF do
     begin
          if (((arrondi (Q1.FindField ('BE_DEBIT').AsFloat, V_PGI.OKDECV)) > 0) or
             ((arrondi(Q1.FindField ('BE_DEBIT').AsFloat,V_PGI.OKDECV)) < 0) and
             ((arrondi (Q1.FindField ('BE_CREDIT').AsFloat,V_PGI.OKDECV)) =0))
             then
                 Sens := 'D'
          else
          if (((Q1.FindField ('BE_CREDIT').AsFloat > 0) or
             (Q1.FindField ('BE_CREDIT').AsFloat < 0)) or (Q1.FindField ('BE_CREDIT').AsFloat>0)) and (Q1.FindField ('BE_DEBIT').AsFloat=0) then
                 Sens := 'C';

            Ligne := Format(FormatmvtBudg, [Q1.FindField('BE_BUDJAL').asstring,
              FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('BE_DATECOMPTABLE').AsDateTime),Q1.FindField('BE_RESOLUTION').asstring,
              Q1.FindField('BE_BUDGENE').asstring, 'B', Q1.FindField('BE_BUDSECT').asstring,
              Q1.FindField('BE_REFINTERNE').asstring, Q1.FindField('BE_LIBELLE').asstring, '', '', Sens,
              QuelMontant(Q1,'BE',0,V_PGI.OkDecV, Sens), Q1.FindField ('BE_QUALIFPIECE').asstring , IntToStr(Q1.FindField ('BE_NUMEROPIECE').asinteger),
              GetParamSocSecur('SO_DEVISEPRINC', 'EUR'), '', '', '', '', Q1.FindField('BE_ETABLISSEMENT').asstring,
              Q1.FindField('BE_AXE').asstring, Q1.FindField('BE_TYPESAISIE').asstring,
              '', '01011900',FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('BE_DATECREATION').AsDateTime),
              '', Q1.FindField('BE_AFFAIRE').asstring, '', Q1.FindField('BE_NATUREBUD').asstring,
              AlignDroite(StrfMontant(Q1.FindField('BE_QTE1').AsFloat,20,4,'',False),20),
              AlignDroite(StrfMontant(Q1.FindField('BE_QTE2').AsFloat,20,4,'',False),20),
              Q1.findField('BE_QUALIFQTE1').Asstring, Q1.findField('BE_QUALIFQTE2').Asstring,

              Q1.findField('BE_LIBRETEXTE1').Asstring, Q1.findField('BE_LIBRETEXTE2').Asstring,
              Q1.findField('BE_LIBRETEXTE3').Asstring, Q1.findField('BE_LIBRETEXTE4').Asstring,
              Q1.findField('BE_LIBRETEXTE5').Asstring,
              Q1.findField('BE_TABLE0').Asstring, Q1.findField('BE_TABLE1').Asstring,
              Q1.findField('BE_TABLE2').Asstring, Q1.findField('BE_TABLE3').Asstring]);
              writeln(F, Ligne);
          Q1.next;
     end;
     Ferme(Q1);
     ExecuteSql('UPDATE BUDECR SET BE_EXPORTE="X" Where '+ WE) ;
end;

procedure  TExport.RenWhereEcr( var Where : string; Ext : string);
var
Jourx                   : string;
EQUALIFPIECE            : string;
Datedeb,Datefin         : TDatetime;
begin
          if (WhereEcriture <> '') and (Ext = 'E') then
          begin
             Where := WhereEcriture; exit;
          end;

          if Trans.Exo <> '' then
             Where := RendCommandeExo (Ext,Trans.Exo);
          if Trans.Jr <> ''  then
          begin
               Jourx := Trans.Jr;
               if Where <> '' then  Where := Where + ' AND ';
               if Ext = 'BE' then
                  Where := Where + '('+ RendCommandeJournalBud(Ext,Jourx)+') '
               else
                   Where := Where + '('+ RendCommandeJournal(Ext,Jourx)+') ';
          end;
         if (Trans.Typearchive = 'JRL') then
         begin
                if BNORMAL  then
                begin
                   EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="N"  ';
                   // fiche 10397
                   if (ExclureEnreg <> '') and (pos('EXO', ExclureEnreg) > 0) and
                   (Ext = 'E')  then
                      EQUALIFPIECE  := EQUALIFPIECE + ' AND E_ECRANOUVEAU<>"OAN" '

                end;
                if BSIMULE  then       // écriture SIMULATION
                begin
                   if EQUALIFPIECE = '' then
                      EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="S"  '
                   else
                      EQUALIFPIECE := EQUALIFPIECE +  'or '+Ext+'_QUALIFPIECE="S" ';
                end;

                if BIFRS  then  //écriture IFRS
                begin
                   if EQUALIFPIECE = '' then
                      EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="I"  '
                   else
                      EQUALIFPIECE := EQUALIFPIECE +  'or '+Ext+'_QUALIFPIECE="I" ';
                end;

                if BPREVISION  then  // écriture prévision
                begin
                   if EQUALIFPIECE = '' then
                      EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="P"  '
                   else
                      EQUALIFPIECE := EQUALIFPIECE +  'or '+Ext+'_QUALIFPIECE="P" ';
                end;

                if BREVISION  then  // écriture révision
                begin
                   if EQUALIFPIECE = '' then
                      EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="R"  '
                   else
                      EQUALIFPIECE := EQUALIFPIECE +  'or '+Ext+'_QUALIFPIECE="R" ';
                end;

                if BSITUATION then // écriture Situation
                begin
                   if EQUALIFPIECE = '' then
                      EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="U"  '
                   else
                   EQUALIFPIECE := EQUALIFPIECE +  'or '+Ext+'_QUALIFPIECE="U" ';
                end;

                if (*(stArg <> '') and*) (TYPEECRX = 'X') then
                begin
                   if EQUALIFPIECE = '' then
                      EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="X"  '
                   else
                   EQUALIFPIECE := EQUALIFPIECE +  'or '+Ext+'_QUALIFPIECE="X" ';
                end;
                if EQUALIFPIECE <> '' then
                begin
                  EQUALIFPIECE := EQUALIFPIECE + ') ';
                  if Where <> '' then  Where := Where + ' AND ';
                     Where := Where + EQUALIFPIECE;
                end;

                if not(BExportStat=cbGrayed) then
                  if BExportStat=cbChecked then Where := Where + ' AND '+Ext+'_EXPORTE="X"' else Where := Where +' AND '+Ext+'_EXPORTE<>"X"' ;

                if (Ext = 'E') and (WhereAvance <> '') then  Where := Where + WhereAvance;
                
                //Date de création  en mode journal
                if (DATEMODIF1<>'') and (DATEMODIF2<>'') and (DATEMODIF1 <> Stdate1900) and  (DATEMODIF2 <> '01/01/2099') then
                begin
                     Datedeb := StrToDate(DATEMODIF1);
                     Datefin := StrToDate(DATEMODIF2);
                     if Datedeb = Datefin then
                     Where := Where + ' AND ('+Ext+'_DATECREATION >="'+ USTIME(StrToDate(DATEMODIF1)) +'"' +
                           ' AND '+Ext+'_DATECREATION < "'+ USTIME(datefin+1) +'")'
                     else
                     Where := Where + ' AND ('+Ext+'_DATECREATION >="'+ USTIME(StrToDate(DATEMODIF1)) +'"' +
                           ' AND '+Ext+'_DATECREATION <="'+ USTIME(StrToDate(DATEMODIF2)) +'")';
                end;
          end
          else
          begin
               if (Trans.Typearchive = 'DOS') then
               begin
                  if (Trans.Serie = 'S1') then
                     EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="N"  '
                     + 'or '+Ext+'_QUALIFPIECE="P" '
                     + 'or '+Ext+'_QUALIFPIECE="R" '
                     + 'or '+Ext+'_QUALIFPIECE="C" '
                     + ' ) '
                   else
                     EQUALIFPIECE := ' ('+Ext+'_QUALIFPIECE="N"  '
                     + 'or '+Ext+'_QUALIFPIECE="S" '
                     + 'or '+Ext+'_QUALIFPIECE="P" '
                     + 'or '+Ext+'_QUALIFPIECE="R" '
                     + 'or '+Ext+'_QUALIFPIECE="C" '
                     + 'or '+Ext+'_QUALIFPIECE="U" ) ';
                   if Where <> '' then  Where := Where + ' AND ';
                     Where := Where + EQUALIFPIECE;
               end;
          end;
          if Trans.Etabi <> '' then
          begin
               if Where <> '' then  Where := Where + ' AND ';
                Where := Where + RendCommandeComboMulti(Ext,Trans.Etabi,'ETABLISSEMENT');
          end;

          if (Trans.Dateecr1 <> iDate1900) and (Trans.Dateecr2 <> iDate1900) then
          begin
               if Where <> '' then  Where := Where + ' AND';
               Where := Where + ' ('+Ext+'_DATECOMPTABLE>="'+USDateTime(Trans.Dateecr1)+'"' +
                              ' AND '+Ext+'_DATECOMPTABLE<="'+USDateTime(Trans.Dateecr2)+'")'
          end
          else
          begin
              if (Trans.Dateecr1 <> iDate1900) then
              begin
                   if Where <> '' then  Where := Where + ' AND';
                   Where := Where + ' ('+Ext+'_DATECOMPTABLE>="'+USDateTime(Trans.Dateecr1)+'")';
              end else
              if (Trans.Dateecr2 <> iDate1900) then
              begin
                   if Where <> '' then  Where := Where + ' AND';
                   Where := Where + ' ('+Ext+'_DATECOMPTABLE<="'+USDateTime(Trans.Dateecr2)+'")';
              end;
          end ;
      if (Trans.Typearchive = 'SYN') and (Ext = 'E') then
      begin
             if (HISTOSYNCHROValue = '0') then
                Where :=  Where + ' AND E_QUALIFPIECE="N" AND E_IO="0"'
             else
             if (HISTOSYNCHROValue = '1') then
                Where :=  Where + ' AND E_QUALIFPIECE="N" AND E_IO="1"'
             else
                Where :=  Where + ' AND E_QUALIFPIECE="N" AND E_IO="X" '
      end;

end;

procedure TExport.ChargementTableLibre(ListeLibre : TList;Where : string);
var
Q1           :TQuery;
TLibre       : PLibre;
typecpte     : string;
begin
Q1:=OpenSQL('SELECT * FROM NATCPTE'+ Where,TRUE) ;
    While Not Q1.Eof Do BEGIN
      System.New (TLibre);
      TLibre^.identifiant := copy (Q1.FindField ('NT_TYPECPTE').asstring,3,1);
      typecpte := copy (Q1.FindField ('NT_TYPECPTE').asstring,1,1);
      TLibre^.code := Q1.FindField ('NT_NATURE').asstring;
      TLibre^.libelle := Q1.FindField ('NT_LIBELLE').asstring;

//T00 à T09 pour les tiers
//G00 à G09 pour les généraux
//S00 à S09 pour les sections
//E00 à E03 pour les ecritures

      if typecpte = 'T' then
         TLibre^.typetable := 'TIE' // pour tiers
      else
          if typecpte = 'G' then
                TLibre^.typetable := 'GEN' // pour généraux
      else
          if typecpte = 'S' then
                TLibre^.typetable := 'SEC' // pour section
      else
          if typecpte = 'E' then
                 TLibre^.typetable := 'ECR' // pour écriture
      else
          if typecpte = 'A' then
                 TLibre^.typetable := 'ANA' // pour analytique
      else
          if typecpte = 'I' then
                 TLibre^.typetable := 'IMM' // pour immo
      else
          if typecpte = 'B' then
                 TLibre^.typetable := 'BUD' // pour budg
      else
          if typecpte = 'D' then
                 TLibre^.typetable := 'DBU' // pour section budg
      else
                 TLibre^.typetable := typecpte; // pour autres
      TLibre^.tx0 := Q1.FindField ('NT_TEXTE0').asstring;
      TLibre^.tx1 := Q1.FindField ('NT_TEXTE1').asstring;
      TLibre^.tx2 := Q1.FindField ('NT_TEXTE2').asstring;
      TLibre^.tx3 := Q1.FindField ('NT_TEXTE3').asstring;
      TLibre^.tx4 := Q1.FindField ('NT_TEXTE4').asstring;
      TLibre^.tx5 := Q1.FindField ('NT_TEXTE5').asstring;
      TLibre^.tx6 := Q1.FindField ('NT_TEXTE6').asstring;
      TLibre^.tx7 := Q1.FindField ('NT_TEXTE7').asstring;
      TLibre^.tx8 := Q1.FindField ('NT_TEXTE8').asstring;
      TLibre^.tx9 := Q1.FindField ('NT_TEXTE9').asstring;
      ListeLibre.Add(TLibre) ;
    Q1.NExt ;
   END ;
Ferme(Q1) ;
end;

procedure TExport.ChargementTableSection(ListeSection : TList;Where : string);
var
Q1,Q2        :TQuery;
TSection     :PSection;
begin
            Q2:=OpenSQL('SELECT PS_SOUSSECTION,PS_LIBELLE,PS_AXE,PS_CODE FROM SSSTRUCR',TRUE) ;
            while not Q2.EOF do
            begin
            System.New (TSection);
            TSection^.code := Q2.FindField ('PS_CODE').asstring;
            TSection^.libelle := Q2.FindField ('PS_LIBELLE').asstring;
            TSection^.axe := Q2.FindField ('PS_AXE').asstring;
            TSection^.codesp := Q2.FindField ('PS_SOUSSECTION').asstring;
                Q1:=OpenSQL('SELECT SS_AXE,SS_SOUSSECTION,SS_LIBELLE,SS_DEBUT,SS_LONGUEUR FROM STRUCRSE Where SS_SOUSSECTION="'+
                                    Q2.FindField ('PS_SOUSSECTION').asstring+'"',TRUE) ;
                if not Q1.EOF then
                begin
                TSection^.libellesp := Q1.FindField ('SS_LIBELLE').asstring;
                TSection^.debsp := Q1.FindField ('SS_DEBUT').asstring;
                TSection^.lgsp := IntToStr(Q1.FindField ('SS_LONGUEUR').asinteger);
                end;
                ferme (Q1);
                ListeSection.Add(TSection) ;
            Q2.NExt ;
            end;
            ferme (Q2);
end;

procedure TExport.Chargementregle(Listeregle : TList;Where : string);
var
Q1                            :TQuery;
TRegle                        :Preglement;
begin
     Q1 := Opensql ('SELECT * from MODEREGL', TRUE);
     while not Q1.EOF do
     begin
          System.new (TRegle);
          TRegle^.code := Q1.FindField ('MR_MODEREGLE').asstring;
          TRegle^.libelle := Q1.FindField ('MR_LIBELLE').asstring;
          TRegle^.apartirde := Q1.FindField ('MR_APARTIRDE').asstring;
          TRegle^.plusjour := Q1.FindField ('MR_PLUSJOUR').asstring;
          TRegle^.arrondijour := Q1.FindField ('MR_ARRONDIJOUR').asstring;
          TRegle^.nbeche := IntTostr(Q1.FindField ('MR_NOMBREECHEANCE').asinteger);
          TRegle^.separepar := Q1.FindField ('MR_SEPAREPAR').asstring;
          TRegle^.montantmin := Q1.FindField ('MR_MONTANTMIN').asstring;
          TRegle^.remplacemin := Q1.FindField ('MR_REMPLACEMIN').asstring;
          TRegle^.mp1 :=  Q1.FindField ('MR_MP1').asstring;
          TRegle^.taux1 := FloatToStr(Q1.FindField('MR_TAUX1').AsFloat);
          TRegle^.mp2 :=  Q1.FindField ('MR_MP2').asstring;
          TRegle^.taux2 := FloatToStr(Q1.FindField('MR_TAUX2').AsFloat);
          TRegle^.mp3 :=  Q1.FindField ('MR_MP3').asstring;
          TRegle^.taux3 := FloatToStr(Q1.FindField('MR_TAUX3').AsFloat);
          TRegle^.mp4 :=  Q1.FindField ('MR_MP4').asstring;
          TRegle^.taux4 := FloatToStr(Q1.FindField('MR_TAUX4').AsFloat);
          TRegle^.mp5 :=  Q1.FindField ('MR_MP5').asstring;
          TRegle^.taux5 := FloatToStr(Q1.FindField('MR_TAUX5').AsFloat);
          TRegle^.mp6 :=  Q1.FindField ('MR_MP6').asstring;
          TRegle^.taux6 := FloatToStr(Q1.FindField('MR_TAUX6').AsFloat);
          TRegle^.mp7 :=  Q1.FindField ('MR_MP7').asstring;
          TRegle^.taux7 := FloatToStr(Q1.FindField('MR_TAUX7').AsFloat);
          TRegle^.mp8 :=  Q1.FindField ('MR_MP8').asstring;
          TRegle^.taux8 := FloatToStr(Q1.FindField('MR_TAUX8').AsFloat);
          TRegle^.mp9 :=  Q1.FindField ('MR_MP9').asstring;
          TRegle^.taux9 := FloatToStr(Q1.FindField('MR_TAUX9').AsFloat);
          TRegle^.mp10 :=  Q1.FindField ('MR_MP10').asstring;
          TRegle^.taux10 := FloatToStr(Q1.FindField('MR_TAUX10').AsFloat);
          TRegle^.mp11 :=  Q1.FindField ('MR_MP11').asstring;
          TRegle^.taux11 :=FloatToStr(Q1.FindField('MR_TAUX11').AsFloat);
          TRegle^.mp12 :=  Q1.FindField ('MR_MP12').asstring;
          TRegle^.taux12 :=FloatToStr(Q1.FindField('MR_TAUX12').AsFloat);
          Listeregle.add (TRegle);
          Q1.next;
      end;
      ferme (Q1);
end;

procedure TExport.ChargementCpte(ListeCpteGen : TList;Where : string; ForcerWhere : Boolean=FALSE);
var
Q1           :TQuery;
TCompteGene  : PListeCpteGen;
//i, j         : integer;
WW,WE,ORDERBY: string;

begin
// i:= gettickcount;
ORDERBY := ' ORDER BY G_GENERAL';
  if (Trans.Typearchive = 'JRL')
  or (Trans.Typearchive = 'SYN') then
  begin
   RenWhereEcr(WE,'E');
   if ForcerWhere then
      WE := Where + ' and g_general not in (select distinct(e_general) from ecriture where '+ WE + ') '
   else
      WE :=' where g_general in (select distinct(e_general) from ecriture where '+ WE + ') ';

   WW :='SELECT G_GENERAL,G_LIBELLE,G_NATUREGENE,G_LETTRABLE,G_POINTABLE,G_VENTILABLE1,'+
   'G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,G_VENTILABLE5,G_TABLE0,'+
   'G_TABLE1,G_TABLE2,G_TABLE3,G_TABLE4,G_TABLE5,G_TABLE6,G_TABLE7,G_TABLE8,G_TABLE9,' +
   'G_ABREGE,G_SENS,G_CORRESP1,G_CORRESP2,G_TVA,G_TVAENCAISSEMENT,G_TPF,G_DATEMODIF,'+
   'G_CUTOFF,G_CUTOFFPERIODE,G_CUTOFFECHUE,G_VISAREVISION,G_CYCLEREVISION,G_CUTOFFCOMPTE, G_CONFIDENTIEL'+
 //Sauvegarde de qualifiant qte  Reprendre dans 6.5
   ',G_QUALIFQTE1, G_QUALIFQTE2 ' +
   ' FROM GENERAUX ' + WE + ORDERBY
  end
  else
   WW :='SELECT G_GENERAL,G_LIBELLE,G_NATUREGENE,G_LETTRABLE,G_POINTABLE,G_VENTILABLE1,'+
   'G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,G_VENTILABLE5,G_TABLE0,'+
   'G_TABLE1,G_TABLE2,G_TABLE3,G_TABLE4,G_TABLE5,G_TABLE6,G_TABLE7,G_TABLE8,G_TABLE9,'+
   'G_ABREGE,G_SENS,G_CORRESP1,G_CORRESP2,G_TVA,G_TVAENCAISSEMENT,G_TPF,G_DATEMODIF,  '+
   'G_CUTOFF,G_CUTOFFPERIODE,G_CUTOFFECHUE,G_VISAREVISION,G_CYCLEREVISION,G_CUTOFFCOMPTE, G_CONFIDENTIEL'+
 //Sauvegarde de qualifiant qte  Reprendre dans 6.5
   ',G_QUALIFQTE1, G_QUALIFQTE2 ' +
   ' FROM GENERAUX'+ Where + ORDERBY;

 Q1:=OpenSQL(WW,TRUE) ;
 While Not Q1.Eof Do BEGIN
      System.New (TCompteGene);
      TCompteGene^.code        := Q1.FindField ('G_GENERAL').asstring;
      TCompteGene^.libelle     := Q1.FindField ('G_LIBELLE').asstring;
      TCompteGene^.nature      := Q1.FindField ('G_NATUREGENE').asstring;
      TCompteGene^.Lettrable   := Q1.FindField ('G_LETTRABLE').asstring;
      TCompteGene^.Pointage    := Q1.FindField ('G_POINTABLE').asstring;
      TCompteGene^.Ventilaxe1  := Q1.FindField ('G_VENTILABLE1').asstring;
      if TCompteGene^.Ventilaxe1 = '' then TCompteGene^.Ventilaxe1 := '-';
      TCompteGene^.Ventilaxe2  := Q1.FindField ('G_VENTILABLE2').asstring;
      if TCompteGene^.Ventilaxe2 = '' then TCompteGene^.Ventilaxe2 := '-';
      TCompteGene^.Ventilaxe3  := Q1.FindField ('G_VENTILABLE3').asstring;
      if TCompteGene^.Ventilaxe3 = '' then TCompteGene^.Ventilaxe3 := '-';
      TCompteGene^.Ventilaxe4  := Q1.FindField ('G_VENTILABLE4').asstring;
      if TCompteGene^.Ventilaxe4 = '' then TCompteGene^.Ventilaxe4 := '-';
      TCompteGene^.Ventilaxe5  := Q1.FindField ('G_VENTILABLE5').asstring;
      if TCompteGene^.Ventilaxe4 = '' then TCompteGene^.Ventilaxe4 := '-';
      TCompteGene^.Table1      := Q1.FindField ('G_TABLE0').asstring;
      TCompteGene^.Table2      := Q1.FindField ('G_TABLE1').asstring;
      TCompteGene^.Table3      := Q1.FindField ('G_TABLE2').asstring;
      TCompteGene^.Table4      := Q1.FindField ('G_TABLE3').asstring;
      TCompteGene^.Table5      := Q1.FindField ('G_TABLE4').asstring;
      TCompteGene^.Table6      := Q1.FindField ('G_TABLE5').asstring;
      TCompteGene^.Table7      := Q1.FindField ('G_TABLE6').asstring;
      TCompteGene^.Table8      := Q1.FindField ('G_TABLE7').asstring;
      TCompteGene^.Table9      := Q1.FindField ('G_TABLE8').asstring;
      TCompteGene^.Table10     := Q1.FindField ('G_TABLE9').asstring;
      TCompteGene^.abrege      := Q1.FindField ('G_ABREGE').asstring;
      TCompteGene^.sens        := Q1.FindField ('G_SENS').asstring;
      TCompteGene^.Corresp1    := Q1.FindField ('G_CORRESP1').asstring;
      TCompteGene^.Corresp2    := Q1.FindField ('G_CORRESP2').asstring;
      TCompteGene^.Tva         := Q1.FindField ('G_TVA').asstring;
      TCompteGene^.Encaissement:= Q1.FindField ('G_TVAENCAISSEMENT').asstring;
      TCompteGene^.TPF         := Q1.FindField ('G_TPF').asstring;
// ajout me 24-05-2005 pour les cutoff

      TCompteGene^.ctoff         := Q1.FindField ('G_CUTOFF').asstring;
      TCompteGene^.ctoffper      := Q1.FindField ('G_CUTOFFPERIODE').asstring;
      TCompteGene^.cteche        := Q1.FindField ('G_CUTOFFECHUE').asstring;
      TCompteGene^.visarev       := Q1.FindField ('G_VISAREVISION').asstring;
      TCompteGene^.cyclerev      := Q1.FindField ('G_CYCLEREVISION').asstring;
      TCompteGene^.ctcpte        := Q1.FindField ('G_CUTOFFCOMPTE').asstring;
      TCompteGene^.condifentiel  := Q1.FindField ('G_CONFIDENTIEL').asstring;
       //Sauvegarde de qualifiant qte  Reprendre dans 6.5
      TCompteGene^.qualqte1  := Q1.FindField ('G_QUALIFQTE1').asstring;
      TCompteGene^.qualqte2  := Q1.FindField ('G_QUALIFQTE2').asstring;
      ListeCpteGen.Add(TCompteGene) ;
    Q1.NExt ;
END ;
Ferme(Q1) ;

// j:= gettickcount;
// showmessage(inttostr(j-i));
end;



FUNCTION TExport.ChargeRib(ListeRib : TList) : Boolean;
var Qr        : Tquery;
TRib          : PListeRib;
WE            : string;
begin
Result := FALSE;
    WE := '';
    if (Trans.Typearchive = 'JRL') or (Trans.Typearchive = 'SYN') then
    begin
     RenWhereEcr(WE,'E');
     WE :=' where R_auxiliaire in (select distinct(e_auxiliaire) from ecriture where '+ WE + ') ';
    end
    else
    begin
         if (Trans.Typearchive = 'BAL') then
            WE :='  Where R_auxiliaire in (select distinct(t_auxiliaire) from tiers '+  CreatWhereParam('T') + ') ';
    end;

    Qr := OpenSql ('SELECT * FROM RIB' + WE, TRUE);
    while not Qr.EOF do
    begin
      System.New (TRib);
       TRib^.ident := 'RIB';
       TRib^.code  := Qr.FindField ('R_AUXILIAIRE').asstring;
       Trib^.numero := IntToStr(Qr.FindField ('R_NUMERORIB').asinteger);
       Trib^.prinpipal := Qr.FindField ('R_PRINCIPAL').asstring;
       Trib^.codebanque :=  Qr.FindField ('R_ETABBQ').asstring;
       Trib^.codeguichet := Qr.FindField ('R_GUICHET').asstring;
       Trib^.nocompte := Qr.FindField ('R_NUMEROCOMPTE').asstring;
       Trib^.cle := Qr.FindField ('R_CLERIB').asstring;
       Trib^.domicile := Qr.FindField ('R_DOMICILIATION').asstring;
       Trib^.ville := Qr.FindField ('R_VILLE').asstring;
       Trib^.pays := Qr.FindField ('R_PAYS').asstring;
       Trib^.devise := Qr.FindField ('R_DEVISE').asstring;
       Trib^.bic := Qr.FindField ('R_CODEBIC').asstring;
       Trib^.soc := Qr.FindField ('R_SOCIETE').asstring;
       Trib^.ribsal := Qr.FindField ('R_SALAIRE').asstring;
       Trib^.ribcompte := Qr.FindField ('R_ACOMPTE').asstring;
       Trib^.ribprof := Qr.FindField ('R_FRAISPROF').asstring;
       Trib^.iban := Qr.FindField ('R_CODEIBAN').asstring;
        // ajout me nature éco et nouveau format des comptes
       Trib^.nateco := Qr.FindField ('R_NATECO').asstring;
       Trib^.typpays := Qr.FindField ('R_TYPEPAYS').asstring;
       ListeRib.Add (Trib);
       Result := TRUE;
       Qr.next;
    end;
ferme (Qr);
end;

procedure TExport.EcritureCpteauxGrandNombre(var F: TextFile);
var
Q1                : TQuery;
Where,WW,MDR      : string;
Ligne             : string;
Compteauxi        : string;
begin
     if Trans.ptiersautre = 'X' then Where := ' ORDER BY T_AUXILIAIRE'
     else
        Where :=  '(t_natureauxi="FOU" or t_natureauxi="CLI" OR t_natureauxi="SAL" or t_natureauxi="DIV"'
        + ' or t_natureauxi="AUC" or t_natureauxi="AUD")'
        + ' ORDER BY T_AUXILIAIRE';
        WW := 'SELECT ##TOP 1000##* FROM TIERS'+ ' Where '+ Where;
        Compteauxi :='0';
      while Compteauxi<>'' do
      begin
         Compteauxi := '';
         Q1:=OpenSQL(WW,TRUE) ;
         While Not Q1.Eof Do BEGIN
         if Q1.FindField ('T_AUXILIAIRE').asstring <> '' then
         begin
              if not CorrepMdr then MDR := Q1.FindField ('T_MODEREGLE').asstring
              else MDR := findcorrespMDR (Q1.FindField ('T_MODEREGLE').asstring, '');
              Ligne := Format(FormatCompteaux, [Q1.FindField ('T_AUXILIAIRE').asstring, Q1.FindField ('T_LIBELLE').asstring,
                Q1.FindField ('T_NATUREAUXI').asstring, Q1.FindField ('T_LETTRABLE').asstring,
                Q1.FindField ('T_COLLECTIF').asstring, Q1.FindField ('T_EAN').asstring, Q1.FindField ('T_TABLE0').asstring,
                Q1.FindField ('T_TABLE1').asstring, Q1.FindField ('T_TABLE2').asstring,
                Q1.FindField ('T_TABLE3').asstring, Q1.FindField ('T_TABLE4').asstring,
                Q1.FindField ('T_TABLE5').asstring, Q1.FindField ('T_TABLE6').asstring,
                Q1.FindField ('T_TABLE7').asstring, Q1.FindField ('T_TABLE8').asstring,
                Q1.FindField ('T_TABLE9').asstring, Q1.FindField ('T_ADRESSE1').asstring,
                Q1.FindField ('T_ADRESSE2').asstring, Q1.FindField ('T_ADRESSE3').asstring,
                Q1.FindField ('T_CODEPOSTAL').asstring, Q1.FindField ('T_VILLE').asstring,
                '', '', '', '', '',
                Q1.FindField ('T_PAYS').asstring, Q1.FindField ('T_ABREGE').asstring, Q1.FindField ('T_LANGUE').asstring,
                Q1.FindField ('T_MULTIDEVISE').asstring, Q1.FindField ('T_DEVISE').asstring, Q1.FindField ('T_TELEPHONE').asstring,
                Q1.FindField ('T_FAX').asstring, Q1.FindField ('T_REGIMETVA').asstring, MDR, Q1.FindField ('T_COMMENTAIRE').asstring,
                Q1.FindField ('T_NIF').asstring, Q1.FindField ('T_SIRET').asstring, Q1.FindField ('T_APE').asstring,
                '', '', '', '', '', '', '', '',
                '', Q1.FindField ('T_FORMEJURIDIQUE').asstring,
                '',Q1.FindField ('T_TVAENCAISSEMENT').asstring,
                // ajout me 18-08-2005
                Q1.FindField ('T_PAYEUR').asstring, Q1.FindField ('T_ISPAYEUR').asstring,
                Q1.FindField ('T_AVOIRRBT').asstring, Q1.FindField ('T_RELANCEREGLEMENT').asstring,
                Q1.FindField ('T_RELANCETRAITE').asstring, Q1.FindField ('T_CONFIDENTIEL').asstring,
                // fiche 10505
                Q1.FindField ('T_CORRESP1').asstring, Q1.FindField ('T_CORRESP2').asstring,
                AlignDroite(StrfMontant(Q1.FindField ('T_ESCOMPTE').asfloat, 20, V_PGI.OkDecV, '', False),20),
                AlignDroite(StrfMontant(Q1.FindField ('T_REMISE').asfloat, 20, V_PGI.OkDecV, '', False),20),
                Q1.FindField ('T_FACTURE').asstring, Q1.FindField ('T_JURIDIQUE').asstring,
                AlignDroite(StrfMontant(Q1.FindField ('T_CREDITDEMANDE').asfloat, 20, V_PGI.OkDecV, '', False),20),
                AlignDroite(StrfMontant(Q1.FindField ('T_CREDITACCORDE').asfloat, 20, V_PGI.OkDecV, '', False),20),
                AlignDroite(StrfMontant(Q1.FindField ('T_CREDITPLAFOND').asfloat, 20, V_PGI.OkDecV, '', False),20)]);
                Compteauxi := Q1.FindField ('T_AUXILIAIRE').asstring;
              writeln(F, Ligne);
          end;
          Q1.NExt ;
         END ;
         WW := 'SELECT ##TOP 1000##* FROM TIERS'+ ' Where T_AUXILIAIRE>"'+Compteauxi+'" AND '+ Where;
         Ferme(Q1) ;
      end;
end;

procedure TExport.ChargementCpteaux(ListeCpteaux : TList; Where : string; ForcerWhere : Boolean=FALSE);
var
Q1          : TQuery;
TCompteAux  : PListeAux;
ii          : integer;
WW,WE       : string;
begin
if (TOBTiers <> nil) and (stArg <> '') then
begin
    for ii :=0 to TOBTiers.detail.Count -1 do
    begin
            System.New (TCompteAux);
            TCompteAux^.code := TOBTiers.Detail[ii].getvalue('T_AUXILIAIRE');
            TCompteAux^.libelle := TOBTiers.Detail[ii].getvalue ('T_LIBELLE');
            TCompteAux.nature := TOBTiers.Detail[ii].getvalue('T_NATUREAUXI');
            TCompteAux^.lettrable  := TOBTiers.Detail[ii].getvalue('T_LETTRABLE');
            TCompteAux^.collectif := TOBTiers.Detail[ii].getvalue ('T_COLLECTIF');
            TCompteAux^.ean := TOBTiers.Detail[ii].getvalue ('T_EAN');
            TCompteAux^.tb1 := TOBTiers.Detail[ii].getvalue ('T_TABLE0');
            TCompteAux^.tb2 := TOBTiers.Detail[ii].getvalue ('T_TABLE1');
            TCompteAux^.tb3 := TOBTiers.Detail[ii].getvalue ('T_TABLE2');
            TCompteAux^.tb4 := TOBTiers.Detail[ii].getvalue ('T_TABLE3');
            TCompteAux^.tb5 := TOBTiers.Detail[ii].getvalue ('T_TABLE4');
            TCompteAux^.tb6 := TOBTiers.Detail[ii].getvalue ('T_TABLE5');
            TCompteAux^.tb7 := TOBTiers.Detail[ii].getvalue ('T_TABLE6');
            TCompteAux^.tb8 := TOBTiers.Detail[ii].getvalue ('T_TABLE7');
            TCompteAux^.tb9 := TOBTiers.Detail[ii].getvalue ('T_TABLE8');
            TCompteAux^.tb10 := TOBTiers.Detail[ii].getvalue ('T_TABLE9');
            TCompteAux^.adresse1 := TOBTiers.Detail[ii].getvalue('T_ADRESSE1');
            TCompteAux^.adresse2 := TOBTiers.Detail[ii].getvalue ('T_ADRESSE2');
            TCompteAux^.adresse3 := TOBTiers.Detail[ii].getvalue ('T_ADRESSE3');
            TCompteAux^.codepostal := TOBTiers.Detail[ii].getvalue ('T_CODEPOSTAL');
            TCompteAux^.ville := TOBTiers.Detail[ii].getvalue ('T_VILLE');
            TCompteAux^.pays := TOBTiers.Detail[ii].getvalue ('T_PAYS');
            TCompteAux^.abrege := TOBTiers.Detail[ii].getvalue ('T_ABREGE');
            TCompteAux^.langue := TOBTiers.Detail[ii].getvalue ('T_LANGUE');
            TCompteAux^.multidevise := TOBTiers.Detail[ii].getvalue ('T_MULTIDEVISE');
            TCompteAux^.devisetiers := TOBTiers.Detail[ii].getvalue ('T_DEVISE');
            TCompteAux^.tel := TOBTiers.Detail[ii].getvalue ('T_TELEPHONE');
            TCompteAux^.fax := TOBTiers.Detail[ii].getvalue('T_FAX');
            TCompteAux^.regimetva := TOBTiers.Detail[ii].getvalue ('T_REGIMETVA');
            TCompteAux^.modereg := TOBTiers.Detail[ii].getvalue('T_MODEREGLE');
            TCompteAux^.commentaire := TOBTiers.Detail[ii].getvalue ('T_COMMENTAIRE');
            TCompteAux^.nif := TOBTiers.Detail[ii].getvalue('T_NIF');
            TCompteAux^.siret := TOBTiers.Detail[ii].getvalue ('T_SIRET');
            TCompteAux^.ape := TOBTiers.Detail[ii].getvalue('T_APE');
            //RechContact(TCompteAux^.code, TCompteAux);
            TCompteAux^.formjur := TOBTiers.Detail[ii].getvalue ('T_FORMEJURIDIQUE');
            TCompteAux^.tvaenc :=  TOBTiers.Detail[ii].getvalue ('T_TVAENCAISSEMENT');
                  // pour les ribs on écrit autant de fois que nombre de rib
            //Okecr := RechRib (TCompteAux^.code, TCompteAux, ListeCpteaux);
            //if not Okecr then
            TCompteAux^.siret := TOBTiers.Detail[ii].getvalue ('T_SIRET');
            TCompteAux^.ape := TOBTiers.Detail[ii].getvalue('T_APE');
            ListeCpteaux.Add(TCompteAux) ;
    end;
    exit;
end;
  if (Trans.Typearchive = 'JRL')
  or (Trans.Typearchive = 'SYN') then
  begin
       RenWhereEcr(WE,'E');
       if ForcerWhere then
          WE := Where + ' and t_auxiliaire not in (select distinct(e_auxiliaire) from ecriture where '+ WE + ') '
       else
          WE :=' where t_auxiliaire in (select distinct(e_auxiliaire) from ecriture where '+ WE + ') ';
       WW :='SELECT * FROM TIERS'+ WE;
  end
  else
  begin
       if Where = '' then
       begin
            if Trans.ptiersautre = 'X' then Where := ' ORDER BY T_AUXILIAIRE'
            else
                Where :=  ' Where  t_natureauxi="FOU" or t_natureauxi="CLI" OR t_natureauxi="SAL" or t_natureauxi="DIV"'
                + ' or t_natureauxi="AUC" or t_natureauxi="AUD"'
                + ' ORDER BY T_AUXILIAIRE';
       end;
       WW := 'SELECT * FROM TIERS'+ Where;
  end;
  Q1:=OpenSQL(WW,TRUE) ;
    While Not Q1.Eof Do BEGIN
      System.New (TCompteAux);
      TCompteAux^.code := Q1.FindField ('T_AUXILIAIRE').asstring;
      TCompteAux^.libelle := Q1.FindField ('T_LIBELLE').asstring;
      TCompteAux.nature := Q1.FindField ('T_NATUREAUXI').asstring;
      TCompteAux^.lettrable  := Q1.FindField ('T_LETTRABLE').asstring;
      TCompteAux^.collectif := Q1.FindField ('T_COLLECTIF').asstring;
      TCompteAux^.ean := Q1.FindField ('T_EAN').asstring;
      TCompteAux^.tb1 := Q1.FindField ('T_TABLE0').asstring;
      TCompteAux^.tb2 := Q1.FindField ('T_TABLE1').asstring;
      TCompteAux^.tb3 := Q1.FindField ('T_TABLE2').asstring;
      TCompteAux^.tb4 := Q1.FindField ('T_TABLE3').asstring;
      TCompteAux^.tb5 := Q1.FindField ('T_TABLE4').asstring;
      TCompteAux^.tb6 := Q1.FindField ('T_TABLE5').asstring;
      TCompteAux^.tb7 := Q1.FindField ('T_TABLE6').asstring;
      TCompteAux^.tb8 := Q1.FindField ('T_TABLE7').asstring;
      TCompteAux^.tb9 := Q1.FindField ('T_TABLE8').asstring;
      TCompteAux^.tb10 := Q1.FindField ('T_TABLE9').asstring;
      TCompteAux^.adresse1 := Q1.FindField ('T_ADRESSE1').asstring;
      TCompteAux^.adresse2 := Q1.FindField ('T_ADRESSE2').asstring;
      TCompteAux^.adresse3 := Q1.FindField ('T_ADRESSE3').asstring;
      TCompteAux^.codepostal := Q1.FindField ('T_CODEPOSTAL').asstring;
      TCompteAux^.ville := Q1.FindField ('T_VILLE').asstring;
      TCompteAux^.pays := Q1.FindField ('T_PAYS').asstring;
      TCompteAux^.abrege := Q1.FindField ('T_ABREGE').asstring;
      TCompteAux^.langue := Q1.FindField ('T_LANGUE').asstring;
      TCompteAux^.multidevise := Q1.FindField ('T_MULTIDEVISE').asstring;
      TCompteAux^.devisetiers := Q1.FindField ('T_DEVISE').asstring;
      TCompteAux^.tel := Q1.FindField ('T_TELEPHONE').asstring;
      TCompteAux^.fax := Q1.FindField ('T_FAX').asstring;
      TCompteAux^.regimetva := Q1.FindField ('T_REGIMETVA').asstring;
      TCompteAux^.modereg := Q1.FindField ('T_MODEREGLE').asstring;
      TCompteAux^.commentaire := Q1.FindField ('T_COMMENTAIRE').asstring;
      TCompteAux^.nif := Q1.FindField ('T_NIF').asstring;
      TCompteAux^.siret := Q1.FindField ('T_SIRET').asstring;
      TCompteAux^.ape := Q1.FindField ('T_APE').asstring;
      TCompteAux^.ctnom := Q1.FindField ('T_PRENOM').asstring;
      //RechContact(TCompteAux^.code, TCompteAux);
      TCompteAux^.formjur := Q1.FindField ('T_FORMEJURIDIQUE').asstring;
      TCompteAux^.tvaenc :=  Q1.FindField ('T_TVAENCAISSEMENT').asstring;
      // ajout me 18-08-2005 fiche 10149
      TCompteAux^.tierspayeur :=  Q1.FindField ('T_PAYEUR').asstring;
      TCompteAux^.ispayeur :=  Q1.FindField ('T_ISPAYEUR').asstring;
      TCompteAux^.remboursement :=  Q1.FindField ('T_AVOIRRBT').asstring;

      // ajout me 26/08/2005
      TCompteAux^.relanceregl  := Q1.FindField ('T_RELANCEREGLEMENT').asstring;
      TCompteAux^.relancetrait := Q1.FindField ('T_RELANCETRAITE').asstring;

      TCompteAux^.confidentiel := Q1.FindField ('T_CONFIDENTIEL').asstring;

      TCompteAux^.corresp1 := Q1.FindField ('T_CORRESP1').asstring;
      TCompteAux^.corresp2 := Q1.FindField ('T_CORRESP2').asstring;
      TCompteAux^.escompte := AlignDroite(StrfMontant(Q1.FindField ('T_ESCOMPTE').asfloat, 20, V_PGI.OkDecV, '', False),20);
      TCompteAux^.remise := AlignDroite(StrfMontant(Q1.FindField ('T_REMISE').asfloat, 20, V_PGI.OkDecV, '', False),20);
      TCompteAux^.facture := Q1.FindField ('T_FACTURE').asstring;
      TCompteAux^.fromjur := Q1.FindField ('T_JURIDIQUE').asstring;
     // fiche com 31027
      TCompteAux^.ceditdemande := AlignDroite(StrfMontant(Q1.FindField ('T_CREDITDEMANDE').asfloat, 20, V_PGI.OkDecV, '', False),20);
      TCompteAux^.ceditaccorde := AlignDroite(StrfMontant(Q1.FindField ('T_CREDITACCORDE').asfloat, 20, V_PGI.OkDecV, '', False),20);
      TCompteAux^.ceditplafond := AlignDroite(StrfMontant(Q1.FindField ('T_CREDITPLAFOND').asfloat, 20, V_PGI.OkDecV, '', False),20);

      ListeCpteaux.Add(TCompteAux) ;
    Q1.NExt ;
   END ;
Ferme(Q1) ;

end;

procedure TExport.EcrTiersCompl(var F: TextFile; Where : string);
var
WE, WW : string;
begin
  RenWhereEcr(WE,'E');
  if (Trans.Serie <> 'S1') then
  begin
        if (Trans.Typearchive = 'JRL') or (Trans.Typearchive = 'SYN') then
         WW :=' where YTC_AUXILIAIRE in (select distinct(e_auxiliaire) from ecriture where '+ WE + ') '
        else
        if (Trans.Typearchive = 'BAL') then  // fiche 10431
        begin
          WE := CreatWhereParam('T');
          WW :=' where YTC_AUXILIAIRE in (select distinct(t_auxiliaire) from tiers '+ WE + ') '
        end
        else
         WW := Where;
        EcritureTierscomp(F, WW);
  end;
  // contact
  if (Trans.Typearchive = 'JRL') or (Trans.Typearchive = 'SYN') then
       WW :=' where C_AUXILIAIRE in (select distinct(e_auxiliaire) from ecriture where '+ WE + ') '
  else
  begin
       WE := CreatWhereParam('T');
       WW :='  Where C_TYPECONTACT="T" AND C_AUXILIAIRE in (select distinct(t_auxiliaire) from tiers '+  WE + ') ';
  end;
  EcritureContact(F, WW);

end;

Procedure TExport.AlimJalB(LJB : TList; Where : string) ;
Var
QJ        : TQuery ;
TJournal  : PListeJournal;
SWhere    : string;
BEGIN
if Where <> '' then SWhere := Where;
QJ:=OpenSQL('SELECT J_JOURNAL,J_LIBELLE,J_NATUREJAL,J_COMPTEURNORMAL,J_COMPTEURSIMUL,'+
 'J_CONTREPARTIE,J_MODESAISIE,J_COMPTEAUTOMAT,J_COMPTEINTERDIT,J_ABREGE,J_AXE FROM JOURNAL'+ sWhere,TRUE) ;
While Not QJ.Eof Do BEGIN
    System.New (TJournal);
    TJournal^.code := QJ.FindField ('J_JOURNAL').asstring;
    TJournal^.libelle := QJ.FindField ('J_LIBELLE').asstring;
    TJournal^.nature := QJ.FindField ('J_NATUREJAL').asstring;
    if (Trans.Serie = 'S1')
    and (TJournal^.nature = 'REG') then TJournal^.nature := 'OD';

    TJournal^.souchen := QJ.FindField ('J_COMPTEURNORMAL').asstring;
    TJournal^.souches := QJ.FindField ('J_COMPTEURSIMUL').asstring;
    TJournal^.compte := QJ.FindField ('J_CONTREPARTIE').asstring;
    TJournal^.axe := QJ.FindField ('J_AXE').asstring;
    TJournal^.modesaisie := QJ.FindField ('J_MODESAISIE').asstring;
    TJournal^.cptauto := QJ.FindField ('J_COMPTEAUTOMAT').asstring;
    TJournal^.cptint := QJ.FindField ('J_COMPTEINTERDIT').asstring;
    TJournal^.abrege := QJ.FindField ('J_ABREGE').asstring;
    LJB.Add(TJournal) ;

    QJ.NExt ;
END ;
if (Trans.Pgene <> 'X') then
if (Trans.balance = TRUE)  and QJ.EOF then
begin
    System.New (TJournal);
    TJournal^.code := 'CO';
    TJournal^.libelle := 'Reprise balance';
    TJournal^.nature := 'OD';
    TJournal^.souchen := 'CPT';
    TJournal^.souches := 'SIM';
    TJournal^.compte := '';
    TJournal^.modesaisie := 'LIB';
    TJournal^.cptauto := '';
    TJournal^.cptint := '';
    TJournal^.abrege := 'Reprise balance';
    LJB.Add(TJournal) ;
    if not (ctxPCL in V_PGI.PGIContexte) then
    begin
              System.New (TJournal);
              TJournal^.code := 'AA';
              TJournal^.libelle := 'Reprise balance Anouveau';
              TJournal^.nature := 'ANO';
              TJournal^.souchen := 'ANO';
              TJournal^.souches := 'SIM';
              TJournal^.compte := '';
              TJournal^.modesaisie := '-';
              TJournal^.cptauto := '';
              TJournal^.cptint := '';
              TJournal^.abrege := 'Reprise balance ANV';
              LJB.Add(TJournal) ;
    end;
end;

Ferme(QJ) ;

END ;

procedure TExport.ChargementMouvement(var ListeMouvt : TList;Where,FormatEnvoie : string;Var F : TextFile; pana : string; ListeJalB : TList);
var
Q1                            :TQuery;
TMvt                          : PListeMouvt;
DecDev,i                      : integer;
Ancdocid                      : string;
CodeMontant                   : string;
EnDevise                      : Boolean ;
Quotite                       : Double;
TDevise                       : TFDevise ;
TDev                          : TList;
Orderby,ExtraitDoc              : string;
Wheres,Verrou,VerrouS5,Whereana : string;
AncJournal,Titre                : string;
Ancexercice, Ancienperiode      : string;
LT9,JointureEcrCompl            : string;
Indice,FCount                   : integer;
TJournal                        : PListeJournal;
TotalMouvDebit, TotalMouvCredit :  double;
Dir1,Ancpiece,SuiteWhere        : string;
ExisteEcrCompl,OkAnvDynamique   : Boolean;
begin

 AncJournal := ''; Indice := 0; Ancdocid := '';
 JointureEcrCompl := '';
 Orderby := ' order by E_JOURNAL, E_EXERCICE, E_PERIODE, E_NUMEROPIECE, E_QUALIFPIECE,E_NUMLIGNE,E_NUMECHE';
 if Where <> '' then
         Wheres := 'Where '+ Where ;
 ExisteEcrCompl := FALSE;
 Fillchar(TJournal, SizeOf(TJournal), #0) ;

 // pour sauvegarder ecrcompl
 if (Trans.Serie <> 'S1')  and ExisteSql ('SELECT EC_EXERCICE FROM ECRCOMPL') then
 begin
    JointureEcrCompl :=  'left outer join ecrcompl on ec_exercice=e_exercice and e_journal=ec_journal '+
    ' and ec_datecomptable=e_datecomptable and ec_numeropiece=e_numeropiece and ec_numligne=e_numligne ';
    ExisteEcrCompl := TRUE;
 end;

 if GerBAP then
 begin
  if JointureEcrCompl = '' then
     JointureEcrCompl :=   'left outer join CPBONSAPAYER on BAP_EXERCICE=e_exercice and BAP_JOURNAL=e_journal ' +
                      ' and BAP_datecomptable=e_datecomptable and BAP_numeropiece=e_numeropiece  '
  else
     JointureEcrCompl := JointureEcrCompl +  'left outer join CPBONSAPAYER on BAP_EXERCICE=e_exercice and BAP_JOURNAL=e_journal ' +
                      ' and BAP_datecomptable=e_datecomptable and BAP_numeropiece=e_numeropiece  '
 end;


 if AvecLettrage then
 begin
     if not OkLettrage (Where, LISTEEXPORT) then
     begin
          AffMessage ('Le lettrage n''est pas équilibré, aucune écriture ne sera exportée');
          exit;
     end;
 end;
 TotalMouvDebit := 0.0;  TotalMouvCredit      :=  0.0;

 //i:= gettickcount;
 DecDev:=V_PGI.OkDecV ;
 Q1:=OpenSQL('SELECT D_DEVISE,D_DECIMALE,D_QUOTITE FROM DEVISE ORDER BY D_DEVISE',True) ;
 TDev:=TList.Create ;
 While not Q1.Eof do
  BEGIN
  TDevise:=TFDevise.Create ;
  TDevise.Code:=Q1.Fields[0].AsString ;
  TDevise.Decimale:=Q1.Fields[1].AsInteger ;
  TDevise.Quotite:=Q1.Fields[2].AsFloat ;
  TDev.Add(TDevise) ;
  Q1.Next ;
  END ;
  Ferme(Q1) ;

  // prendre dans la table devise le decimal pour les arrondi

if ListeJalB <> nil then
begin
   FCount := ListeJalB.Count-1;
   if (FCount < 0) and (NATUREJRL <>'ODA') then
      OnAfficheListeCom('Les critères de sélection ne renvoient aucun mouvement.',LISTEEXPORT);
end
else
   FCount := 0; // correction format sage en export dedoublement des écritures dans le fichier
{$IFDEF SCANGED}
if BGed then
begin
     Dir1             := ExtractFileDir(Trans.FichierSortie) +'\GED';
     if DirectoryExists(Dir1) then
       RemoveInDir2 (Dir1, TRUE, TRUE, TRUE);
     CreateDir(Dir1);
end;
{$ENDIF}

// AJOUT SUR 7XX
// Pour ANouveau dynamique, il ne faut pas envoyé les anouveaux
if (Trans.Serie = 'S1') and GetParamSocSecur('SO_CPANODYNA',false)
and (GetEncours.EtatCpta = 'CPR') then
begin
    OkAnvDynamique := TRUE;  SuiteWhere := ' AND (E_EXERCICE <> "'+ GetSuivant.Code +'")';
end
else
    OkAnvDynamique := FALSE;

for i := 0 to FCount do
begin

       if ListeJalB <> nil then
       begin
          TJournal := ListeJalB.Items[i];
          if Where <> '' then
           Wheres := 'Where E_JOURNAL="'+TJournal^.code+'" AND ('+ Where +')';
       end;
       // AJOUT SUR 7XX
       if OkAnvDynamique and (TJournal^.code = GetParamSocSecur('SO_JALOUVRE', 'ANO'))then
          Wheres := Wheres + SuiteWhere;

       Q1:=OpenSQL('SELECT * FROM ECRITURE '+ JointureEcrCompl + Wheres + OrderBy, TRUE);

       While Not Q1.Eof Do BEGIN
              Application.ProcessMessages;
              if Indice > 100 then
              begin
                         EcritureMouvement(ListeMouvt,F,Trans.TypeFormat,CorrepMdp);
                         ListeMouvt.Free ;
                         ListeMouvt:=TList.Create;
                         Indice := 0;
              end;
        // écriture des dans le fichier dès qu'il y a une rupture E_JOURNAL, E_EXERCICE, E_PERIODE, E_NUMEROPIECE
              if AncJournal <> Q1.FindField ('E_JOURNAL').asstring then
              begin
                   EcritureMouvement(ListeMouvt,F,Trans.TypeFormat,CorrepMdp);
                   ListeMouvt.Free ;
                   ListeMouvt:=TList.Create;
                   OnAfficheListeCom('Ecriture du journal : '+ Q1.FindField ('E_JOURNAL').asstring,LISTEEXPORT);
              end;
              AncJournal := Q1.FindField ('E_JOURNAL').asstring;

              if Ancexercice <> Q1.FindField ('E_EXERCICE').asstring then
              begin
                   EcritureMouvement(ListeMouvt,F,Trans.TypeFormat,CorrepMdp);
                   ListeMouvt.Free ;
                   ListeMouvt:=TList.Create;
              end;
              Ancexercice := Q1.FindField ('E_EXERCICE').asstring;

              if (AncienPeriode <> Q1.FindField ('E_PERIODE').asstring) then
              begin
                   EcritureMouvement(ListeMouvt,F,Trans.TypeFormat,CorrepMdp);
                   ListeMouvt.Free ;
                   ListeMouvt:=TList.Create;
              end;
              AncienPeriode := Q1.FindField ('E_PERIODE').asstring;

              System.New (TMvt);
              TMvt^.Journal := Q1.FindField ('E_JOURNAL').asstring;
              if Trans.Serie = 'SAGE' then
                 TMvt^.Datecomptable := FormatDateTime(Traduitdateformat('ddmmyy'),Q1.FindField ('E_DATECOMPTABLE').AsDateTime)
              else
                 TMvt^.Datecomptable := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('E_DATECOMPTABLE').AsDateTime);
              TMvt^.Typepiece := Q1.FindField ('E_NATUREPIECE').asstring;

              if (Trans.Serie = 'SAGE') then
              begin
                 if TMvt^.Typepiece='OC' then TMvt^.Typepiece:='RC' else if TMvt^.Typepiece='OF' then TMvt^.Typepiece:='RF' ;
              end;

              TMvt^.General := Q1.FindField ('E_GENERAL').asstring;
              TMvt^.AuxSect := Q1.FindField ('E_AUXILIAIRE').asstring;
              if TMvt^.AuxSect <>'' then TMvt^.Typecpte:= 'X' else TMvt^.Typecpte:=' ' ;
              TMvt^.Refinterne := Q1.FindField ('E_REFINTERNE').asstring;
              TMvt^.Libelle := Q1.FindField ('E_LIBELLE').asstring;
              TMvt^.Modepaie := Q1.FindField ('E_MODEPAIE').asstring;
             // if (Trans.Serie = 'S1') and (TMvt^.Modepaie <> '') then CoherenceMDRS1(TMvt^.Modepaie) ;

              if Trans.Serie = 'SAGE' then
              begin
                 TMvt^.Modepaie :=CHERCHEUNMODE(TMvt^.Modepaie) ;
                 TMvt^.Echeance := FormatDateTime(Traduitdateformat('ddmmyy'),Q1.FindField ('E_DATEECHEANCE').AsDateTime);
              end
              else
                 TMvt^.Echeance := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('E_DATEECHEANCE').AsDateTime);

              if (((arrondi (Q1.FindField ('E_DEBIT').AsFloat, V_PGI.OKDECV)) > 0) or
                 ((arrondi(Q1.FindField ('E_DEBIT').AsFloat,V_PGI.OKDECV)) < 0) and
                 ((arrondi (Q1.FindField ('E_CREDIT').AsFloat,V_PGI.OKDECV)) =0))
              then
                   TMvt^.Sens := 'D';

              if (((Q1.FindField ('E_CREDIT').AsFloat > 0) or
              (Q1.FindField ('E_CREDIT').AsFloat < 0)) or (Q1.FindField ('E_CREDIT').AsFloat>0)) and (Q1.FindField ('E_DEBIT').AsFloat=0) then
                   TMvt^.Sens := 'C';
              // pour la mode
              if (*(stArg <> '') and EComsx ne marche pas*) (TYPEECRX = 'X') and (ECRINTEGRE <> '') then
                 TMvt^.Typeecriture := ECRINTEGRE
              else
                 TMvt^.Typeecriture := Q1.FindField ('E_QUALIFPIECE').asstring;
              TMvt^.NumPiece := IntToStr(Q1.FindField ('E_NUMEROPIECE').asinteger);
              TMvt^.Devise := Q1.FindField ('E_DEVISE').asstring;
              TMvt^.TauxDev := AlignDroite(StrfMontant(Q1.FindField ('E_COTATION').asfloat,20,9,'',False),10);

              CodeMontant := 'E--';
              EnDevise:=RecupDevise(TMvt^.Devise,DecDev,Quotite,TDev) ;
              If EnDevise Then
              BEGIN
                              CodeMontant:='DE-' ;
                              TMvt^.Montant1:=QuelMontant(Q1,'E',1,DecDev,TMvt^.Sens) ;
                              TMvt^.Montant2:=QuelMontant(Q1,'E',0,V_PGI.OkDecV,TMvt^.Sens) ;
                              TMvt^.Montant3:=Format_String(' ',20) ;
              END Else
              BEGIN
                              CodeMontant:='E--';
                              TMvt^.Montant1:=QuelMontant(Q1,'E',0,V_PGI.OkDecV,TMvt^.Sens) ;
                              TMvt^.Montant2:=Format_String(' ',20) ;
                              TMvt^.Montant3:=Format_String(' ',20) ;
             END ;

           TMvt^.CodeMontant := Codemontant;

           if TMvt^.Sens = 'C' then
              TotalMouvCredit := TotalMouvCredit + Q1.Findfield('E_CREDIT').AsFloat
           else
              TotalMouvDebit  := TotalMouvDebit  + Q1.Findfield('E_DEBIT').AsFloat;

           if (EtablissUnique) then
              TMvt^.Etablissement := GetParamSocSecur ('SO_ETABLISDEFAUT', '')
           else
              TMvt^.Etablissement := Q1.FindField ('E_ETABLISSEMENT').asstring;
           TMvt^.Axe := '';

           TMvt^.Numeche := IntToStr(Q1.FindField ('E_NUMECHE').asinteger);

             if FormatEnvoie = 'ETE' then // format étendu
             begin
                 TMvt^.RefExterne := Q1.findField('E_REFEXTERNE').Asstring;
                 TMvt^.Daterefexterne := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('E_DATEREFEXTERNE').AsDateTime);
                 TMvt^.Datecreation := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('E_DATECREATION').AsDateTime);
                 TMvt^.Societe := Q1.findField('E_SOCIETE').Asstring;
                 TMvt^.Affaire := Q1.findField('E_AFFAIRE').Asstring;
                 TMvt^.Datetauxdev := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('E_DATETAUXDEV').AsDateTime);
                 TMvt^.Ecranouveau := Q1.findField('E_ECRANOUVEAU').Asstring;
                 TMvt^.Qte1 := AlignDroite(StrfMontant(Q1.FindField('E_QTE1').AsFloat,20,4,'',False),20);
                 TMvt^.Qte2 := AlignDroite(StrfMontant(Q1.FindField('E_QTE2').AsFloat,20,4,'',False),20);
                 TMvt^.QualifQte1 := Q1.findField('E_QUALIFQTE1').Asstring;
                 TMvt^.QualifQte2 := Q1.findField('E_QUALIFQTE2').Asstring;
                 TMvt^.Reflibre := Q1.findField('E_REFLIBRE').Asstring;
                 TMvt^.Tvaencaiss := Q1.findField('E_TVAENCAISSEMENT').Asstring;
                 TMvt^.Regimetva := Q1.findField('E_REGIMETVA').Asstring;
                 TMvt^.Tva := Q1.findField('E_TVA').Asstring;
                 TMvt^.TPF := Q1.findField('E_TPF').Asstring;
                 TMvt^.Contrepartigen := Q1.findField('E_CONTREPARTIEGEN').Asstring;
                 TMvt^.contrepartiaux := Q1.findField('E_CONTREPARTIEAUX').Asstring;
                 TMvt^.RefPointage := Q1.findField('E_REFPOINTAGE').Asstring;
                 TMvt^.datepointage := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('E_DATEPOINTAGE').AsDateTime);
                 TMvt^.daterelance := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('E_DATERELANCE').AsDateTime);
                 TMvt^.datevaleur :=  FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('E_DATEVALEUR').AsDateTime);
                 TMvt^.Rib := Q1.findField('E_RIB').Asstring;
                 TMvt^.Refreleve := Q1.findField('E_REFRELEVE').Asstring;
                 if (Trans.Typearchive <> 'JRL') then
                    TMvt^.Numeroimmo := Q1.findField('E_IMMO').Asstring
                 else
                    TMvt^.Numeroimmo := '';
                 TMvt^.LT0 := Q1.findField('E_LIBRETEXTE0').Asstring;
                 TMvt^.LT1 := Q1.findField('E_LIBRETEXTE1').Asstring;
                 TMvt^.LT2 := Q1.findField('E_LIBRETEXTE2').Asstring;
                 TMvt^.LT3 := Q1.findField('E_LIBRETEXTE3').Asstring;
                 TMvt^.LT4 := Q1.findField('E_LIBRETEXTE4').Asstring;
                 TMvt^.LT5 := Q1.findField('E_LIBRETEXTE5').Asstring;
                 TMvt^.LT6 := Q1.findField('E_LIBRETEXTE6').Asstring;
                 TMvt^.LT7 := Q1.findField('E_LIBRETEXTE7').Asstring;
                 TMvt^.LT8 := Q1.findField('E_LIBRETEXTE8').Asstring;
                 if (Trans.Serie = 'S1') or (ctxPCL in V_PGI.PGIContexte) then
                 begin
                       if Q1.findField('E_ETATREVISION').Asstring = 'X' then
                           Verrou := 'X'
                       else
                           Verrou := '-';
                       if (Q1.findField('E_PAQUETREVISION').Asinteger = 1) and (Q1.findField('E_LETTRAGE').asstring <> '') then
                           VerrouS5 := 'X'
                 end
                 else
                 begin
                      if not (ctxPCL in V_PGI.PGIContexte) then
                      begin
                           if Q1.findField('E_REFGESCOM').Asstring  <> '' then Verrou := 'X'
                           {JP 26/06/07 : FQ TRESO 10491 : on vérouille les flux originaires de la Tréso}
                           else if Q1.findField('E_QUALIFORIGINE').AsString  = QUALIFTRESO then Verrou := 'X'
                           else  Verrou := '-';
                      end;
                 end;
                 LT9 := '';
                 if ((ctxPCL in V_PGI.PGIContexte) and ((Trans.Typearchive = 'SYN') or ((Trans.Serie = 'S1') and (Trans.Typearchive = 'DOS')) or
                 ((Trans.Serie <> 'S1') and (Trans.Typearchive = 'DOSS')))) then
                 begin
                              if (Q1.findField('E_REFREVISION').Asinteger = 0)
                              or ((Trans.Serie = 'S1') and (Trans.Typearchive = 'DOS'))
                              or ((Trans.Serie <> 'S1') and (Trans.Typearchive = 'DOSS') )
                              then
                              begin
                                if (Trans.Typearchive = 'SYN') then
                                 TMvt^.LT9 := Format ('      %1.1s%-3.3s%-6.6s%-.06d%1.1s%-4.4s%d',
                                 [Verrou, TMvt^.Journal, Q1.findField('E_PERIODE').Asstring, Q1.FindField ('E_NUMEROPIECE').asinteger, VerrouS5,
                                 Copy (Q1.findField('E_LIBRETEXTE9').Asstring,24,4),Q1.findField('E_PAQUETREVISION').Asinteger])
                                 else

                                 TMvt^.LT9 := Format ('      %1.1s%-3.3s%-6.6s%-.06d%1.1s%-4.4s',
                                 [Verrou, TMvt^.Journal, Q1.findField('E_PERIODE').Asstring, Q1.FindField ('E_NUMEROPIECE').asinteger, VerrouS5,
                                 Copy (Q1.findField('E_LIBRETEXTE9').Asstring,24,4)])
                              end
                              else
                                  TMvt^.LT9 := Format ('%.06d%1.1s%-3.3s%-6.6s%-.06d%1.1s%-4.4s%d',
                                  [Q1.findField('E_REFREVISION').Asinteger, Verrou, TMvt^.Journal,
                                  Q1.findField('E_PERIODE').Asstring, Q1.FindField ('E_NUMEROPIECE').asinteger, VerrouS5,
                                  Copy (Q1.findField('E_LIBRETEXTE9').Asstring,24,4),Q1.findField('E_PAQUETREVISION').Asinteger]);
                              LT9 := TMvt^.LT9;
                 end
                 else
                 if (Not (ctxPCL in V_PGI.PGIContexte) and ((Trans.Typearchive = 'SYN') or (Trans.Typearchive = 'DOSS'))) then
                 begin
                              if (Q1.findField('E_REFREVISION').Asinteger = 0)
                              or ((Trans.Serie <> 'S1') and (Trans.Typearchive = 'DOSS'))
                              then
                              begin
                                if (Trans.Typearchive = 'SYN') then
                                 TMvt^.LT9 := Format ('%6d%1.1s%-3.3s%-6.6s           %d',
                                 [Q1.FindField ('E_NUMEROPIECE').asinteger, Verrou, TMvt^.Journal, Q1.findField('E_PERIODE').Asstring,Q1.findField('E_PAQUETREVISION').Asinteger])
                                 else
                                 TMvt^.LT9 := Format ('%6d%1.1s%-3.3s%-6.6s           ',
                                 [Q1.FindField ('E_NUMEROPIECE').asinteger, Verrou, TMvt^.Journal, Q1.findField('E_PERIODE').Asstring])
                              end
                              else
                                  TMvt^.LT9 := Format ('%6d%1.1s%-3.3s%-6.6s%-.06d%1.1s%-4.4s%d',
                                  [Q1.FindField ('E_NUMEROPIECE').asinteger, Verrou, TMvt^.Journal,
                                  Q1.findField('E_PERIODE').Asstring, Q1.findField('E_REFREVISION').Asinteger, VerrouS5,
                                  Copy (Q1.findField('E_LIBRETEXTE9').Asstring,24,4),Q1.findField('E_PAQUETREVISION').Asinteger]);
                              LT9 := TMvt^.LT9;
                 end
                 else
                     TMvt^.LT9 := Q1.findField('E_LIBRETEXTE9').Asstring;

                 TMvt^.TA0 := Q1.findField('E_TABLE0').Asstring;
                 TMvt^.TA1 := Q1.findField('E_TABLE1').Asstring;
                 TMvt^.TA2 := Q1.findField('E_TABLE2').Asstring;
                 TMvt^.TA3 := Q1.findField('E_TABLE3').Asstring;
                 TMvt^.LM0 := Q1.findField('E_LIBREMONTANT0').Asstring;
                 TMvt^.LM1 := Q1.findField('E_LIBREMONTANT1').Asstring;
                 TMvt^.LM2 := Q1.findField('E_LIBREMONTANT2').Asstring;
                 TMvt^.LM3 := Q1.findField('E_LIBREMONTANT3').Asstring;
                 TMvt^.LD  := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('E_LIBREDATE').AsDateTime);
                 TMvt^.LB0 := Q1.findField('E_LIBREBOOL0').Asstring;
                 TMvt^.LB1 := Q1.findField('E_LIBREBOOL1').Asstring;
                 TMvt^.Conso := Q1.findField('E_CONSO').Asstring;
                 TMvt^.Datepaquetmax :=  FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('E_DATEPAQUETMAX').AsDateTime);
                 TMvt^.Datepaquetmin :=  FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('E_DATEPAQUETMIN').AsDateTime);
                 // on enleve le lettrage dans le cas du journal
                 TMvt^.lettragelibre := '-';
                 // ajout me 01/03/2004
                 if (Trans.Typearchive = 'JRL') and (not AvecLettrage) then
                 begin
                     TMvt^.Couverture    := AlignDroite(StrfMontant(0,20,V_PGI.OkDecV,'',False),20);
                     TMvt^.Couverturedev := AlignDroite(StrfMontant(0,20,V_PGI.OkDecV,'',False),20);
                     TMvt^.Couverturelibre :=AlignDroite(StrfMontant(0,20,V_PGI.OkDecE,'',False),20);
                     TMvt^.lettrage := '';
                     TMvt^.lettragedev := '-';
                     if Q1.findField('E_ETATLETTRAGE').Asstring = 'RI' then
                        TMvt^.etatlettrage := 'RI'
                     else
                        TMvt^.etatlettrage := 'AL';

                     TMvt^.RefPointage := '';
                     TMvt^.datepointage := FormatDateTime(Traduitdateformat('ddmmyyyy'),iDate1900);
                 end
                 else
                 begin
                     TMvt^.Couverture := AlignDroite(StrfMontant(Q1.FindField('E_COUVERTURE').AsFloat,20,V_PGI.OkDecV,'',False),20);
                     TMvt^.Couverturedev := AlignDroite(StrfMontant(Q1.FindField('E_COUVERTUREDEV').AsFloat,20,V_PGI.OkDecV,'',False),20);
                     TMvt^.Couverturelibre := AlignDroite(StrfMontant(0,20,V_PGI.OkDecE,'',False),20);
                     TMvt^.lettrage := Q1.findField('E_LETTRAGE').Asstring;
                     TMvt^.lettragedev := Q1.findField('E_LETTRAGEDEV').Asstring;
                     TMvt^.etatlettrage := Q1.findField('E_ETATLETTRAGE').Asstring;
                 end;
                 TMvt^.refgescom  := Q1.findField('E_REFGESCOM').Asstring;
                 TMvt^.typemvt    := Q1.findField('E_TYPEMVT').asstring;
                 TMvt^.Treso      := Q1.findField('E_TRESOSYNCHRO').asstring;
                 TMvt^.numtrait   := Q1.findField('E_NUMTRAITECHQ').asstring;
                 TMvt^.numenca    := Q1.findField('E_NUMENCADECA').asstring;
                 TMvt^.valide     := Q1.findField('E_VALIDE').asstring;
                 if AnnulValidation and (TMvt^.valide = 'X') then TMvt^.valide := '-';
                 // fiche 14216
                 TMvt^.confidentiel := Q1.FindField ('E_CONFIDENTIEL').asstring;
                 // ajout me 07/12/2005
                 TMvt^.cfonbok      := Q1.FindField ('E_CFONBOK').asstring;
                 TMvt^.codeaccept   := Q1.FindField ('E_CODEACCEPT').asstring;
                 TMvt^.qualiforigine:= Q1.FindField ('E_QUALIFORIGINE').asstring;

                 // ajout me 26-05-2005 pour ecrcompl
                 if ExisteEcrCompl then
                 begin
                      if Q1.FindField ('EC_JOURNAL').asstring <> '' then
                      begin
                           TMvt^.ctdeb      := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('EC_CUTOFFDEB').AsDateTime);
                           TMvt^.ctfin      := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('EC_CUTOFFFIN').AsDateTime);
                           TMvt^.ctdate     := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('EC_CUTOFFDATECALC').AsDateTime);
                           TMvt^.cleecr     := Q1.findField('EC_CLEECR').asstring;
                           TMvt^.docid      := Q1.findField('EC_DOCGUID').asstring;
                      end;
                 end;
             end;
           ListeMouvt.Add(TMvt) ;
           Inc(Indice);
           if (Q1.FindField('E_ANA').AsString='X') And (Pana = 'X' ) then // traitement analytique
           begin
                Whereana := 'Y_JOURNAL="'+TMvt^.Journal+'" AND Y_DATECOMPTABLE="'+
                USDateTime (Q1.findField('E_DATECOMPTABLE').AsDatetime) +
                '" AND Y_NUMEROPIECE=' + TMvt^.NumPiece +
                ' AND Y_NUMLIGNE='+ IntToStr(Q1.findField('E_NUMLIGNE').Asinteger);
                ChargementMouvementAnalytique(ListeMouvt, Whereana, FormatEnvoie, Tdev, F, Indice, LT9);
//                ExecuteSql('UPDATE ANALYTIQ SET Y_EXPORTE="X" Where '+Whereana) ;
           end;
           // Document GED
           if BGed and (Trans.Serie <> 'SAGE') and (TMvt^.docid <> '')  and (JointureEcrCompl <> '') then
           begin
              if Ancdocid <> Q1.findField('EC_DOCGUID').asstring then
              begin
                  System.New (TMvt);
                  TMvt^.Journal    := Q1.FindField ('E_JOURNAL').asstring;
                  TMvt^.Datecomptable := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('E_DATECOMPTABLE').AsDateTime);
                  TMvt^.Typepiece  := Q1.FindField ('E_NATUREPIECE').asstring;
                  TMvt^.General    := Q1.FindField ('E_GENERAL').asstring;
                  TMvt^.AuxSect    := Q1.FindField ('E_AUXILIAIRE').asstring;
                  TMvt^.Typecpte   := 'G';
                  TMvt^.Refinterne := Q1.FindField ('E_REFINTERNE').asstring;
                  TMvt^.docid      := ExtraitDocument (Q1.findField('EC_DOCGUID').asstring, Titre, FALSE);
                  Copyfile (PChar(TMvt^.docid), Pchar(Dir1+'\'+ExtractFileName(TMvt^.docid)),TRUE);
{$IFNDEF EAGLSERVER}
                  TMvt^.docid      :='.\GED\'+ExtractFileName(TMvt^.docid);
{$ELSE}
                  TMvt^.docid      :='.\'+ExtractFileName(TMvt^.docid);
{$ENDIF}
                  ListeMouvt.Add(TMvt) ;
                  Ancdocid         := Q1.findField('EC_DOCGUID').asstring;
              end;
           end;
           if GerBAP then
           begin
              if (Q1.FindField ('BAP_JOURNAL').asstring <> '') and
              (Ancpiece <> Q1.FindField ('E_NUMEROPIECE').asstring) then
              begin
                begin
                  System.New (TMvt);
                  TMvt^.Journal    := Q1.FindField ('BAP_JOURNAL').asstring;
                  TMvt^.Datecomptable := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('BAP_DATECOMPTABLE').AsDateTime);
                  TMvt^.Typecpte   := 'P';
                  TMvt^.NumPiece   := IntToStr(Q1.FindField ('E_NUMEROPIECE').asinteger);
                  TMvt^.viseur     := Q1.FindField('BAP_VISEUR').asstring;
                  TMvt^.statutbap  := Q1.FindField('BAP_STATUTBAP').asstring;
                  TMvt^.viseur1    := Q1.FindField('BAP_VISEUR1').asstring;
                  TMvt^.viseur2    := Q1.FindField('BAP_VISEUR2').asstring;
                  TMvt^.numeroordre    := IntToStr(Q1.FindField('BAP_NUMEROORDRE').asinteger);
                  TMvt^.Nbjour     := IntToStr(Q1.FindField('BAP_NBJOUR').asinteger);
                  TMvt^.Echeance   := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField('BAP_ECHEANCEBAP').AsDateTime);
                  TMvt^.tierspayeur:= Q1.FindField('BAP_TIERSPAYEUR').asstring;
                  TMvt^.cleecr     := Q1.FindField('BAP_CLEFFACTURE').asstring;
                  TMvt^.codevisa   := Q1.FindField('BAP_CODEVISA').asstring;
                  TMvt^.circuitbap := Q1.FindField('BAP_CIRCUITBAP').asstring;
                  TMvt^.ctdate     := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField('BAP_DATEMAIL').AsDateTime);
                  TMvt^.daterelance  := FormatDateTime(Traduitdateformat('ddmmyyyy'), Q1.FindField('BAP_DATERELANCE1').AsDateTime);
                  TMvt^.daterelance2 := FormatDateTime(Traduitdateformat('ddmmyyyy'), Q1.FindField('BAP_DATERELANCE2').AsDateTime);
                  TMvt^.DateEcheance := FormatDateTime(Traduitdateformat('ddmmyyyy'), Q1.FindField('BAP_DATEECHEANCE').AsDateTime);
                  TMvt^.societe      := Q1.FindField('BAP_SOCIETE').asstring;
                  TMvt^.Etablissement:= Q1.FindField('BAP_ETABLISSEMENT').asstring;
                  TMvt^.alerte1      := Q1.FindField('BAP_ALERTE1').asstring;
                  TMvt^.alerte2      := Q1.FindField('BAP_ALERTE2').asstring;
                  TMvt^.relance1     := Q1.FindField('BAP_RELANCE1').asstring;
                  TMvt^.relance2     :=Q1.FindField('BAP_RELANCE2').asstring;
                  ListeMouvt.Add(TMvt) ;
                end;
              end;
              Ancpiece := Q1.FindField ('E_NUMEROPIECE').asstring;
           end;
           Q1.NExt ;
        END ;
        Ferme(Q1) ;
        if ListeJalB <> nil then
        begin
           EcritureMouvement(ListeMouvt,F,Trans.TypeFormat,CorrepMdp);
           ListeMouvt.free;
           ListeMouvt:=TList.Create;
        end;
        ExecuteSql('UPDATE ECRITURE SET E_EXPORTE="X" '+Wheres) ;

end;  // For

// pour les ODA
if (Trans.Typearchive = 'JRL') and (NATUREJRL ='ODA') then
begin
//SH_ANALYTIQUE= 'X'
        RenWhereEcr(Whereana,'Y');
        ChargementMouvementAnalytique(ListeMouvt, Whereana, FormatEnvoie, Tdev, F, Indice, LT9);
end;
// pour les OD analytiques a voir od analytique pour synchro
if (Trans.Serie <> 'S1') and (Trans.Typearchive = 'SYN') then
begin
     Whereana := '';
     RenWhereEcr(Whereana,'Y');
     if Whereana <> '' then Whereana := Whereana + ' AND Y_TYPEANALYTIQUE="X"'
     else Whereana := 'Where Y_TYPEANALYTIQUE="X"';
     ChargementMouvementAnalytique(ListeMouvt,Whereana,FormatEnvoie, Tdev, F, Indice);
end;
if Where <> '' then Wheres := 'Where '+ Where;

if TDev <> nil then begin VideListe(TDev); TDev.Free ; end;

if TotalMouvDebit <> 0 then
 OnAfficheListeCom(TraduireMemoire('Total mouvements débiteurs  : '+ StrfMontant(TotalMouvDebit, 20, V_PGI.OkDecV,'',False)), LISTEEXPORT);
if TotalMouvCredit <> 0 then
 OnAfficheListeCom(TraduireMemoire('Total mouvements créditeurs : '+ StrfMontant(TotalMouvCredit, 20, V_PGI.OkDecV,'',False)), LISTEEXPORT);
if (FCount >= 0) and (NATUREJRL <>'ODA') and (TotalMouvDebit = 0)
and (TotalMouvDebit = 0)then
      OnAfficheListeCom('Les critères de sélection ne renvoient aucun mouvement.',LISTEEXPORT);

end;

procedure TExport.ChargementBalance(ListeMouvt : TList; Var F : TextFile);
var
TMvt                          : PListeMouvt;
Jrl                           : string;
MttSolde                      : double;
Compte                        : string;
Exo                           : string;
nature,libelle, Collectif     : string;
OkAuxi                        : Boolean;
Q1                            : TQuery;
SQL                           : string;
EAUXILIAIRE                   : string;
EQUALIFPIECE,Etabliss         : string;
begin

 OkAuxi := AUXILIARE;
 Jrl := ReadTokenSt(Trans.Jr);
//GG if (Jrl = '') or (Jrl = '<<Tous>>')  then
 if (Jrl = '') or (Jrl = Traduirememoire('<<Tous>>'))  then
 begin
     Trans.Jr := 'CO';
     Jrl := 'CO';
 end;

 if Trans.Exo <> '' then Exo := Trans.Exo
 else Exo := GetEncours.Code;
 if Trans.Dateecr1 = iDate1900 then Trans.Dateecr1 := GetEncours.Deb;
 if (Trans.Dateecr2 = iDate1900) or (DateToStr(Trans.Dateecr2)='01/01/2099') then Trans.Dateecr2 := GetEncours.Fin;

 if Okauxi then  EAUXILIAIRE := 'E_AUXILIAIRE,'
 else EAUXILIAIRE := ' ';

 if BNORMAL  then
    EQUALIFPIECE := ' (E_QUALIFPIECE="N"  ' ;
 if BSIMULE then
 begin
    if EQUALIFPIECE = '' then
       EQUALIFPIECE := ' (E_QUALIFPIECE="S"  '
    else
       EQUALIFPIECE := EQUALIFPIECE +  'or E_QUALIFPIECE="S" ';
 end;
 if BSITUATION then
 begin
    if EQUALIFPIECE = '' then
       EQUALIFPIECE := ' (E_QUALIFPIECE="U"  '
    else
       EQUALIFPIECE := EQUALIFPIECE +  'or E_QUALIFPIECE="U" ';
 end;
 if (stArg <> '') and (TYPEECRX = 'X') then
 begin
     if EQUALIFPIECE = '' then
        EQUALIFPIECE := ' ('+'E_QUALIFPIECE="X"  '
     else
     EQUALIFPIECE := EQUALIFPIECE +  'or '+'E_QUALIFPIECE="X" ';
 end;
 if EQUALIFPIECE <> '' then
       EQUALIFPIECE := ' AND '+ EQUALIFPIECE + ') ';

 if Trans.Etabi <> '' then
 //Etabliss := 'AND (E_ETABLISSEMENT="'+ Trans.Etabi+'") '
   Etabliss := 'AND ('+ RendCommandeComboMulti('E',Trans.Etabi,'ETABLISSEMENT')+') '
 else Etabliss := '';


 Sql := 'SELECT E_GENERAL,'+ EAUXILIAIRE + 'E_EXERCICE, ' +
     'sum(E_DEBIT) TOTDEBIT, sum(E_CREDIT) TOTCREDIT, E_QUALIFPIECE,'+
     ' G_NATUREGENE,G_LIBELLE FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' +
     'WHERE  E_DATECOMPTABLE>="'+USDateTime(Trans.Dateecr1)+'" AND E_DATECOMPTABLE<="'+USDateTime(Trans.Dateecr2) +
     '" AND E_EXERCICE="'+Exo+'" '+ EQUALIFPIECE +
     Etabliss +
     ' GROUP BY E_GENERAL,'+ EAUXILIAIRE+ 'E_EXERCICE,E_QUALIFPIECE,G_NATUREGENE,G_LIBELLE '+
     ' ORDER BY E_GENERAL,'+ EAUXILIAIRE+'E_EXERCICE,E_QUALIFPIECE,G_NATUREGENE,G_LIBELLE';
 Q1 := OpenSql (Sql, TRUE);

 MttSolde := 0.0;
 while not Q1.EOF do
 begin
      Compte  := Q1.FindField ('E_GENERAL').asstring;
      Nature  := Q1.FindField ('G_NATUREGENE').asstring;
      Libelle := Q1.FindField ('G_LIBELLE').asstring;
      if Okauxi then
      Collectif  := Q1.FindField ('E_AUXILIAIRE').asstring
      else
      Collectif := '';

      if Collectif <> '' then
         OnAfficheListeCom('Compte : ' + Collectif, LISTEEXPORT)
      else
          OnAfficheListeCom('Compte : ' + Compte, LISTEEXPORT);
      MttSolde := arrondi (Q1.FindField ('TOTDEBIT').asFloat - Q1.FindField ('TOTCREDIT').asFloat,V_PGI.OKDECV);
      if MttSolde <> 0.0 then
      begin
                 System.New (TMvt);
                 TMvt^.Journal := Jrl;
                 RemplirEcritureBal (TMvt, Compte, Collectif, nature, libelle, MttSolde);
                 ListeMouvt.Add(TMvt) ;
                 // pour analytiq
                if BANA then  // pour ventilation analytique
                begin
                     Traitement_Ventil (TMvt, ListeMouvt,Exo,Compte, Collectif, nature, libelle, Jrl);
                end;

      end;
      Q1.next;
    END ;
    if (ListeMouvt <> nil) and (ListeMouvt.Count > 0 ) then
            ListeMouvt.Sort(CompareJournal);
    Ferme (Q1);
end;



FUNCTION TExport.CHERCHEUNMODE (Mode : String3) : Char ;
BEGIN
      if Mode='ESP'           then ChercheUnMode:='E' else
      if Mode='CHQ'           then ChercheUnMode:='C' else
      if Mode='CB'            then ChercheUnMode:='U' else
      if Mode = 'CTB'         then ChercheUnMode:='U' else
      if Mode='VIR'           then ChercheUnMode:='V' else
      if Mode='PRE'           then ChercheUnMode:='P' else
      if Mode='TRD'           then ChercheUnMode:='T' else
      if Mode='TRA'           then ChercheUnMode:='T' else
      if Mode='LCR'           then ChercheUnMode:='L' else
      if Mode='BOR'           then ChercheUnMode:='B' else
      if Mode='LCR'           then ChercheUnMode:='T' else
      ChercheUnMode:='S' ;
END ;
                                                                                                                                               
procedure TExport.ChargementMouvementAnalytique(var ListeMouvt : TList;Where,FormatEnvoie : string; Tdev : TList;Var F : TextFile; Var Indice : integer; LT9: string ='');
var
Q1                            :TQuery;
TMvt                          : PListeMouvt;
DecDev                        : integer ;
CodeMontant                   : string;
EnDevise                      : Boolean ;
Quotite                       : Double;
Orderby                       : string;
Wheres,lStSQL                 : string;
begin

 Orderby := ' ORDER BY Y_JOURNAL,Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_QUALIFPIECE,Y_AXE, Y_NUMVENTIL';
 if Where <> '' then
 begin
         Wheres := 'Where '+ Where;// + ' AND Y_QUALIFPIECE="N"';
 end;
 DecDev:=V_PGI.OkDecV ;

  // prendre dans la table devise le decimal pour les arrondi

if FormatEnvoie = 'STD' then
Q1:=OpenSQL('SELECT Y_JOURNAL,Y_DATECOMPTABLE,Y_NATUREPIECE, Y_GENERAL,Y_QUALIFPIECE,'+
'Y_SECTION,Y_REFINTERNE,Y_LIBELLE,Y_DEBIT,Y_CREDIT,'+
'Y_NUMEROPIECE,Y_DEVISE,Y_TAUXDEV,Y_ETABLISSEMENT,Y_DEBIT,Y_CREDIT,'+
'Y_DEBITDEV,Y_CREDITDEV, Y_AXE, Y_TYPEANALYTIQUE,Y_CONFIDENTIEL FROM ANALYTIQ '+ Wheres + Orderby,TRUE)
else
begin
     lStSQL:=CGetSQLFromTable('ANALYTIQ',['Y_BLOCNOTE'] ) ; // construit le texte de la requete sql sans les champs passe en parametre
     Q1:=OpenSQL(lStSQL+' '+ Wheres + OrderBy,TRUE);
end;

While Not Q1.Eof Do BEGIN
      System.New (TMvt);

      TMvt^.Journal := Q1.FindField ('Y_JOURNAL').asstring;
      if Trans.Serie = 'SAGE' then
         TMvt^.Datecomptable := FormatDateTime(Traduitdateformat('ddmmyy'),Q1.FindField ('Y_DATECOMPTABLE').AsDateTime)
      else
         TMvt^.Datecomptable := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('Y_DATECOMPTABLE').AsDateTime);
      TMvt^.Typepiece := Q1.FindField ('Y_NATUREPIECE').asstring;
      TMvt^.General := Q1.FindField ('Y_GENERAL').asstring;
      TMvt^.AuxSect := Q1.FindField ('Y_SECTION').asstring;
      if Q1.FindField ('Y_TYPEANALYTIQUE').asstring = 'X' then
         TMvt^.Typecpte:='O'
      else
         TMvt^.Typecpte:='A' ;
      TMvt^.Refinterne := Q1.FindField ('Y_REFINTERNE').asstring;
      TMvt^.Libelle := Q1.FindField ('Y_LIBELLE').asstring;
      TMvt^.Echeance := '        ';

      if (((arrondi (Q1.FindField ('Y_DEBIT').AsFloat, V_PGI.OKDECV)) > 0) or
         ((arrondi(Q1.FindField ('Y_DEBIT').AsFloat,V_PGI.OKDECV)) < 0) and
         ((arrondi (Q1.FindField ('Y_CREDIT').AsFloat,V_PGI.OKDECV)) =0))
      then
           TMvt^.Sens := 'D';

      if (((Q1.FindField ('Y_CREDIT').AsFloat > 0) or
      (Q1.FindField ('Y_CREDIT').AsFloat < 0)) or (Q1.FindField ('Y_CREDIT').AsFloat>0)) and (Q1.FindField ('Y_DEBIT').AsFloat=0) then
           TMvt^.Sens := 'C';

      if (*(stArg <> '') and EComsx ne marche pas*) (TYPEECRX = 'X') and (ECRINTEGRE <> '') then
         TMvt^.Typeecriture := ECRINTEGRE
      else
         TMvt^.Typeecriture := Q1.FindField ('Y_QUALIFPIECE').asstring;
      TMvt^.NumPiece := IntToStr(Q1.FindField ('Y_NUMEROPIECE').asinteger);
      TMvt^.Devise := Q1.FindField ('Y_DEVISE').asstring;
      TMvt^.TauxDev := AlignDroite(StrfMontant(Q1.FindField ('Y_TAUXDEV').asfloat,20,9,'',False),10);

      CodeMontant := 'E--';
      EnDevise:=RecupDevise(TMvt^.Devise,DecDev,Quotite,TDev) ;
      If EnDevise Then
      BEGIN
           CodeMontant:='DE-' ;
           TMvt^.Montant1:=QuelMontant(Q1,'Y',1,DecDev,TMvt^.Sens) ;
           TMvt^.Montant2:=QuelMontant(Q1,'Y',0,V_PGI.OkDecV,TMvt^.Sens) ;
           TMvt^.Montant3:=Format_String(' ',20) ;
      END Else
      BEGIN
                      CodeMontant:='E--';
                      TMvt^.Montant1:=QuelMontant(Q1,'Y',0,V_PGI.OkDecV,TMvt^.Sens) ;
                      TMvt^.Montant2:=Format_String(' ',20) ;
                      TMvt^.Montant3:=Format_String(' ',20) ;
     END ;
   TMvt^.CodeMontant := Codemontant;
   if (EtablissUnique)  then
      TMvt^.Etablissement := GetParamSocSecur ('SO_ETABLISDEFAUT', '')
   else
      TMvt^.Etablissement := Q1.FindField ('Y_ETABLISSEMENT').asstring;
   TMvt^.Axe := Q1.FindField ('Y_AXE').asstring;
   if FormatEnvoie = 'ETE' then // format étendu
   begin
         TMvt^.RefExterne := Q1.findField('Y_REFEXTERNE').Asstring;
         TMvt^.Daterefexterne := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('Y_DATEREFEXTERNE').AsDateTime);
         TMvt^.Datecreation := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('Y_DATECREATION').AsDateTime);
         TMvt^.Societe := Q1.findField('Y_SOCIETE').Asstring;
         TMvt^.Affaire := Q1.findField('Y_AFFAIRE').Asstring;
         TMvt^.Datetauxdev := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('Y_DATETAUXDEV').AsDateTime);
         TMvt^.Ecranouveau := Q1.findField('Y_ECRANOUVEAU').Asstring;
         TMvt^.Qte1 := AlignDroite(StrfMontant(Q1.FindField('Y_QTE1').AsFloat,20,4,'',False),20);
         TMvt^.Qte2 := AlignDroite(StrfMontant(Q1.FindField('Y_QTE2').AsFloat,20,4,'',False),20);
         TMvt^.QualifQte1 := Q1.findField('Y_QUALIFQTE1').Asstring;
         TMvt^.QualifQte2 := Q1.findField('Y_QUALIFQTE2').Asstring;
         TMvt^.Reflibre := Q1.findField('Y_REFLIBRE').Asstring;
         TMvt^.Contrepartigen := Q1.findField('Y_CONTREPARTIEGEN').Asstring;
         TMvt^.contrepartiaux := Q1.findField('Y_CONTREPARTIEAUX').Asstring;
         TMvt^.LT0 := Q1.findField('Y_LIBRETEXTE0').Asstring;
         TMvt^.LT1 := Q1.findField('Y_LIBRETEXTE1').Asstring;
         TMvt^.LT2 := Q1.findField('Y_LIBRETEXTE2').Asstring;
         TMvt^.LT3 := Q1.findField('Y_LIBRETEXTE3').Asstring;
         TMvt^.LT4 := Q1.findField('Y_LIBRETEXTE4').Asstring;
         TMvt^.LT5 := Q1.findField('Y_LIBRETEXTE5').Asstring;
         TMvt^.LT6 := Q1.findField('Y_LIBRETEXTE6').Asstring;
         TMvt^.LT7 := Q1.findField('Y_LIBRETEXTE7').Asstring;
         TMvt^.LT8 := Q1.findField('Y_LIBRETEXTE8').Asstring;
         if LT9 <> '' then  
            TMvt^.LT9 := LT9
         else
            TMvt^.LT9 := Q1.findField('Y_LIBRETEXTE9').Asstring;
         TMvt^.TA0    := Q1.findField('Y_TABLE0').Asstring;
         TMvt^.TA1    := Q1.findField('Y_TABLE1').Asstring;
         TMvt^.TA2    := Q1.findField('Y_TABLE2').Asstring;
         TMvt^.TA3    := Q1.findField('Y_TABLE3').Asstring;
         TMvt^.LM0    := Q1.findField('Y_LIBREMONTANT0').Asstring;
         TMvt^.LM1    := Q1.findField('Y_LIBREMONTANT1').Asstring;
         TMvt^.LM2    := Q1.findField('Y_LIBREMONTANT2').Asstring;
         TMvt^.LM3    := Q1.findField('Y_LIBREMONTANT3').Asstring;
         TMvt^.LD     := FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('Y_LIBREDATE').AsDateTime);
         TMvt^.LB0    := Q1.findField('Y_LIBREBOOL0').Asstring;
         TMvt^.LB1    := Q1.findField('Y_LIBREBOOL1').Asstring;
         TMvt^.Conso  := Q1.findField('Y_CONSO').Asstring;
         TMvt^.typemvt:= Q1.findField('Y_TYPEMVT').asstring;
         // axes croisés  V7
         TMvt^.souspl1    := Q1.findField('Y_SOUSPLAN1').Asstring;
         TMvt^.souspl2    := Q1.findField('Y_SOUSPLAN2').Asstring;
         TMvt^.souspl3    := Q1.findField('Y_SOUSPLAN3').Asstring;
         TMvt^.souspl4    := Q1.findField('Y_SOUSPLAN4').Asstring;
         TMvt^.souspl5    := Q1.findField('Y_SOUSPLAN5').Asstring;
         TMvt^.valide     := Q1.findField('Y_VALIDE').asstring;
         if AnnulValidation and (TMvt^.valide = 'X') then TMvt^.valide := '-';
         // fiche 14216
         TMvt^.confidentiel := Q1.FindField ('Y_CONFIDENTIEL').asstring;

     end;
   ListeMouvt.Add(TMvt) ;
   Inc(Indice);
   Q1.NExt ;
END ;
Ferme(Q1) ;
if Where <> '' then Wheres := 'Where '+ Where;
end;

procedure TExport.RemplirEcritureBal (var TMvt : PListeMouvt; Compte, comptecol, nature, libelle : string; var MttSolde : double);
var
CodeMontant,Tmp                   : string;
begin
     TMvt^.Datecomptable := FormatDateTime(Traduitdateformat('ddmmyyyy'),Trans.Dateecr2);

     TMvt^.Typepiece := 'OD';
     TMvt^.General := Compte;
     TMvt^.AuxSect := comptecol;

     if (nature = 'COF') and (TMvt^.AuxSect = '') then  TMvt^.AuxSect := GetParamSocSecur ('SO_FOUATTEND', '');
     if (nature = 'COC') and (TMvt^.AuxSect = '') then TMvt^.AuxSect := GetParamSocSecur ('SO_CLIATTEND', '');
     if (nature = 'COD') and (TMvt^.AuxSect = '') then TMvt^.AuxSect := GetParamSocSecur ('SO_DIVATTEND', '');
     if (nature = 'COS') and (TMvt^.AuxSect = '') then TMvt^.AuxSect := GetParamSocSecur ('SO_SALATTEND', '');

     if TMvt^.AuxSect <> '' then TMvt^.Typecpte:= 'X' else TMvt^.Typecpte:=' ' ;
     TMvt^.Libelle := libelle;
     TMvt^.Echeance := FormatDateTime(Traduitdateformat('ddmmyyyy'),iDate1900);
     TMvt^.Typeecriture := 'N';
     TMvt^.Devise := V_PGI.DevisePivot;
     TMvt^.TauxDev := '1';
     if TMvt^.Devise = 'FRF' then   CodeMontant := 'F--';
     if TMvt^.Devise = 'EUR' then   CodeMontant := 'E--';
     TMvt^.CodeMontant := Codemontant;
     Tmp := Trans.Etabi;
     TMvt^.Etablissement := ReadTokenSt (Tmp);
//     TMvt^.Etablissement := Trans.Etabi;
     if TMvt^.Etablissement = '' then
        TMvt^.Etablissement := GetParamSocSecur ('SO_ETABLISDEFAUT', '');
     TMvt^.Axe := '';
     TMvt^.Montant2:=Format_String(' ',20) ;
     TMvt^.Montant3:=Format_String(' ',20) ;
     if (arrondi (MttSolde, V_PGI.OKDECV)) > 0 then TMvt^.Sens := 'D';
     if (arrondi (MttSolde, V_PGI.OKDECV)) < 0 then TMvt^.Sens := 'C';
     if MttSolde < 0.0 then MttSolde := MttSolde * (-1);
     TMvt^.Montant1 := StrfMontant(MttSolde,20,V_PGI.OkDecV,'',False);
end;

procedure TExport.Traitement_Ventil (var TMvt : PListeMouvt; var ListeMouvt : TList; Exo,Compte, Collectif, nature, libelle, Jrl : string);
var
MttSolde : double;
SQl      : string;
Q2       : TQuery;
begin
     Sql := 'SELECT Y_GENERAL, Y_AUXILIAIRE, Y_EXERCICE, ' +
     'sum(Y_DEBIT) TOTDEBIT, sum(Y_CREDIT) TOTCREDIT, Y_SECTION, Y_AXE '+
     'FROM ANALYTIQ ' +
     'WHERE  Y_DATECOMPTABLE>="'+USDateTime(Trans.Dateecr1)+'" AND Y_DATECOMPTABLE<="'+USDateTime(Trans.Dateecr2) +
     '" AND Y_EXERCICE="'+Exo +   '" AND Y_ECRANOUVEAU="N"  AND Y_GENERAL="' + Compte + '"'+
     ' GROUP BY Y_GENERAL,Y_AUXILIAIRE,Y_EXERCICE, Y_SECTION, Y_AXE '+
     ' ORDER BY Y_GENERAL,Y_AUXILIAIRE,Y_EXERCICE, Y_SECTION, Y_AXE';
     MttSolde := 0.0;
     Q2 := OpenSql (Sql, TRUE);
     while not Q2.EOF do
     begin
          MttSolde := arrondi (Q2.FindField ('TOTDEBIT').asFloat - Q2.FindField ('TOTCREDIT').asFloat, V_PGI.OKDECV);
          if MttSolde <> 0.0 then
          begin
            System.New (TMvt);
            TMvt^.Journal := Jrl;
            RemplirEcritureBal (TMvt, Compte, Collectif, nature, libelle, MttSolde);
            TMvt^.AuxSect := Q2.FindField ('Y_SECTION').asstring;
            TMvt^.axe     := Q2.FindField ('Y_AXE').asstring;
            TMvt^.Typecpte:= 'A';
            ListeMouvt.Add(TMvt) ;
          end;
          Q2.next;
     end;
     ferme (Q2);

end;

{$IFDEF EAGLSERVER}
Function TExport.AfficheComExport(Chaine: string; Listecom : TListBox) : Boolean;
begin
     if Listecom = nil then
     begin
          writeln(FicIE, Chaine);
          cWA.MessagesAuClient('COMSX.EXPORT','',Chaine) ;
     end;
     Result := TRUE;
end;
{$ENDIF}

procedure TExport.ChargementParamlibre(Where : string; Var TobParam : TOB );
var
Q1           :TQuery;
T1           : TOB;
begin
 if TobParam = nil then
    TobParam := TOB.create('Input_TOBM', nil, -1);

 Q1 := Opensql('SELECT * FROM PARAMLIB left join CHOIXCOD on CC_CODE=PL_TABLE ' + Where + ' ORDER BY PL_TABLE', TRUE);
 While Not Q1.Eof Do BEGIN
      T1 := TobParam.FindFirst(['PL_TABLE'],  [Q1.FindField ('PL_TABLE').asstring], FALSE);
      if T1 = nil then
      begin
            T1 := TOB.Create ('',TobParam,-1);
            T1.AddChampSupValeur('LIBELLE', Q1.FindField ('CC_LIBELLE').asstring);
            if Q1.FindField ('CC_ABREGE').asstring = 'X' then
            T1.AddChampSupValeur('VISIBLE', Q1.FindField ('CC_ABREGE').asstring)
            else
            T1.AddChampSupValeur('VISIBLE', '-');
            T1.AddChampSupValeur('PL_TABLE', Q1.FindField ('PL_TABLE').asstring);
      end;
      T1.AddChampSupValeur(Q1.FindField ('PL_CHAMP').asstring,  Q1.FindField ('PL_LIBELLE').asstring);
      T1.AddChampSupValeur('V'+Q1.FindField ('PL_CHAMP').asstring,  Q1.FindField ('PL_VISIBLE').asstring);
    Q1.NExt ;
 END ;
Ferme(Q1) ;

end;

{$IFDEF EAGLSERVER}
const toSQL  = 9; 
function TExport.ExportBobEagl (Fich : string; var TB : TOB) : Boolean;
var
Ind            : integer;
TSql           : TOB;
Q1             : TQuery;
TableName,Requete      : string;
begin
            for Ind := 0 to Tb.detail.count-1 do
            begin
                 TSql := Tb.detail[Ind];
                 if (TSql.GetValue ('_OBJECTNAME') <> '') then
                 begin
                               Requete := TSql.GetValue ('_OBJECTNAME');
                               Q1 := OpenSql(Requete, TRUE);
                               //TData := TOB.Create('', TSql, -1);
                               TableName := Requete;
                               Requete := ReadTokenPipe (TableName, 'FROM ');
                               if Pos(' ', TableName) = 0  then
                                  TableName  := Copy(TableName, 0, length (TableName))
                               else
                                   TableName  := Copy(TableName, 0, Pos(' ', TableName));

                               TSql.LoadDetailDB(TableName, '', '', Q1, TRUE, FALSE);
                               Ferme(Q1);
                 end;
            end;
        TB.SaveToBinFile (Fich, FALSE, TRUE, FALSE, TRUE);
        Result := TRUE;
end;
{$ENDIF}

Procedure TExport.ExportLesIMMOS;
var
TIMMO : TOB;
begin
    if not ExisteSQL('SELECT * FROM IMMO') then exit;

    TIMMO := TOB.Create('Maman',nil,-1);
    try
      RenseignelaSerie(ExeCOMSX) ;
      TIMMO.AddChampSupValeur('BOBNAME', Trans.FichierSortie);
      TIMMO.AddChampSupValeur('BOBVERSION', 7);
      TIMMO.AddChampSupValeur('BOBNUMSOCREF', V_PGI.NumVersionBase);
      TIMMO.AddChampSupValeur('BOBDATEGEN', date);
//      TIMMO.AddChampSupValeur('BOBPRODUIT', 'IMMO');
      OnAfficheListeCom('Export des Immobilisations ', LISTEEXPORT);
      AddToBob ( TIMMO, 'SELECT * FROM IMMO');
      OnAfficheListeCom('Export des amortissements ', LISTEEXPORT);
      AddToBob ( TIMMO, 'SELECT * FROM IMMOAMOR');
      OnAfficheListeCom('Export des comptes ', LISTEEXPORT);
      AddToBob ( TIMMO, 'SELECT * FROM IMMOCPTE');
      OnAfficheListeCom('Export des paramettres sociétés ', LISTEEXPORT);
      AddToBob ( TIMMO, 'SELECT * FROM PARAMSOC WHERE SOC_TREE LIKE "001;002%"');

      // tablette coefficient dégressif
      OnAfficheListeCom('Export des tables libres ', LISTEEXPORT);
      AddToBob ( TIMMO, 'select * FROM choixcod where CC_TYPE="ICD"');
      AddToBob ( TIMMO, 'select * FROM choixcod where CC_TYPE="GEO"');
      AddToBob ( TIMMO, 'select * FROM choixcod where CC_TYPE="MDC"');
      AddToBob ( TIMMO, 'select * FROM natcpte  Where NT_TYPECPTE like "I0%"');
      AddToBob ( TIMMO, 'SELECT * FROM  CHOIXCOD where CC_CODE like "I0%"');
      if AglExportBob ( Trans.FichierSortie, False,False, TIMMO,True) then
         OnAfficheListeCom('Export des immobilisations terminé ', LISTEEXPORT)
      else
         OnAfficheListeCom('Erreur lors de l''exportation des immobilisations ', LISTEEXPORT);

    finally
      TIMMO.Free;
    end;
end;

{$IFDEF EAGLSERVER}

procedure TExport.FileClickzip(FFile : string='');
var
  FileName       : String ;
  Commentaire    : String ;
  TheToz         : TOZ;
  Password       : string;
  Filearchive    : string;
  FFi            : string;
 procedure AfficheErreur (msg : string);
 begin
        OnAfficheListeCom ( 'Erreur;compression répertoire dat, soit la session n''est pas ouverte en ajout.', LISTEEXPORT ) ;
        TheToz.CancelSession ;
 end;
begin
  Password := '';
    // Récupération du nom du fichier a insérer
    //
  FileName := Trans.FichierSortie;
  Filearchive := ReadTokenPipe (Filename, '.');
  Filearchive := Filearchive + '.ZIP';

  TheToz := TOZ.Create ;
  try
    if TheToz.OpenZipFile ( Filearchive, moCreate ) then
    begin
        if TheToz.OpenSession ( osAdd ) then
        begin
            if FFile = '' then
            begin
               FFi := Trans.FichierSortie;
               if FileExists(FFi) then
               begin
                     if not TheToz.ProcessFile ( FFi, Commentaire ) then
                     begin
                        OnAfficheListeCom ( 'Erreur;Soit le fichier : ' + ExtractFileName ( FileName ) + ' n''existe plus, soit la session n''est pas ouverte en ajout.', LISTEEXPORT ) ;
                        TheToz.CancelSession ;
                     end ;
               end
               else
                        TheToz.CancelSession ;
            end
            else
            begin
               if FFile = '$DOS' then // répertoire DOS
               begin
                  if Not ZippeDirectory(TheToz, PathDos, '*.*', True, False, '.MDF;.LDF;.mdf;ldf') then
                     AfficheErreur ('Erreur;compression répertoire dossier, soit la session n''est pas ouverte en ajout.')
               end
               else
               begin
                       if FFile = '$STD' then // répertoire STD
                       begin
                          if Not ZippeDirectory(TheToz, PathStd, '*.*', True, False, '.MDF;.LDF;.mdf;ldf') then
                             AfficheErreur ('Erreur;compression répertoire standard, soit la session n''est pas ouverte en ajout.')
                       end
                       else
                       begin
                       if FFile = '$DAT' then // répertoire DAT
                       begin
                          if Not ZippeDirectory(TheToz, PathDat, '*.*', True, False, '.MDF;.LDF;.mdf;ldf') then
                              AfficheErreur ('Erreur;compression répertoire dat, soit la session n''est pas ouverte en ajout.') ;
                       end
                       else
                       if FFile = '$DB0' then  // sauvegarde annuaire etc... de DB0
                       begin
{$IFDEF EAGLSERVER}
                           if Not LanceMainSauveDossier (TheToz, ExtractFileDir(FFi), Nodossier, 'S') then
                              AfficheErreur ('Erreur sauvegarde de la base cabinet, dossier permanent');
{$ENDIF}
                       end
                       else
                       begin
                          FFi := FFile;
                          if Not ZippeDirectory(TheToz, FFI, '*.*', True, False) then
                              AfficheErreur ( 'Erreur;compression répertoire dat, soit la session n''est pas ouverte en ajout.') ;
                       end;
                       end;
               end;
            end;
        end
        else
        OnAfficheListeCom ( 'Erreur;Soit le fichier : ' + ExtractFileName ( FileName ) + ' n''existe plus, soit la session n''est pas ouverte en ajout.', LISTEEXPORT ) ;
    end
    else
    begin
      OnAfficheListeCom ( 'Erreur création du fichier archive : ' + 'archive.zip' + ' impossible', LISTEEXPORT) ;
      Exit ;
    end ;
 EXCEPT
    On E: Exception do
      begin
      OnAfficheListeCom ( 'TozError : ' + E.Message, LISTEEXPORT ) ;
      end ;
 END ;
  TheToz.CloseSession ;
  if FFile = '' then DeleteFile(Trans.FichierSortie);
  TheToz.Free;
 Trans.FichierSortie := Filearchive;
end;
{$ENDIF}

procedure TExport.AddToBob(TStd: TOB; stSQL : string; sDomaine: string='C');
begin
  with TOB.Create('',TStd,-1) do
  begin
    AddChampSupValeur('_OBJECTTYPE', toSQL);
    AddChampSupValeur('_OBJECTNAME', stSQL);
    AddChampSupValeur('_OBJECTWITHDATA', False );
    AddChampSupValeur('_OBJECTDOMAINE', sDomaine);
  end;
end;

Procedure TExport.ExportLaCompta;
var
TCPTA : TOB;
TSql  : TOB;
Q1    : TQuery;
i     : integer;
begin              
    TSql :=TOB.Create('',Nil,-1) ;
    Q1 := OpenSql ('select * from detables where (dt_domaine="C") or' +
    ' DT_NOMTABLE="EXERCICE"', TRUE);
    TSql.LoadDetailDB('detables', '', '', Q1, TRUE, FALSE);
    Ferme(Q1);

    TCPTA := TOB.Create('Maman',nil,-1);
    try
      RenseignelaSerie(ExeCOMSX) ;
      TCPTA.AddChampSupValeur('BOBNAME', Trans.FichierSortie);
      TCPTA.AddChampSupValeur('BOBVERSION', 7);
      TCPTA.AddChampSupValeur('BOBNUMSOCREF', V_PGI.NumVersionBase);
      TCPTA.AddChampSupValeur('BOBDATEGEN', date);
      For i := 0  to TSql.detail.count-1 do
      begin
          OnAfficheListeCom(TSql.detail[i].getvalue ('DT_LIBELLE'), LISTEEXPORT);
          AddToBob ( TCPTA, 'SELECT * FROM '+ TSql.detail[i].getvalue ('DT_NOMTABLE'), TSql.detail[i].getvalue ('DT_DOMAINE'));
      end;
      if AglExportBob (Trans.FichierSortie, False, False, TCPTA, True) then
        OnAfficheListeCom('Export des informations comptables terminé ', LISTEEXPORT)
      else
         OnAfficheListeCom('Erreur lors de l''exportation des informations comptables ', LISTEEXPORT);

    finally
      TCPTA.Free;
    end;
end;


end.
