unit UtilMetres;

interface
uses {$IFDEF VER150} variants,{$ENDIF}
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Hctrls, ExtCtrls, StdCtrls, Mask,HEnt1,HMsgBox,EntGC,
  HSysMenu, Menus,lookup,FileCtrl,comobj,ActiveX, CBPPath,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  Doc_Parser,DBCtrls, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main,
{$ENDIF}
{$IFDEF BTP}
  BTPUTIL,
  UtilFichiers,
  UTofBtSaisieVar,
  BTVARDOC_TOF,
{$ENDIF}
  DicoBTP,UTOB, AffaireUtil,EtudesUtil,EtudesExt,SaisUtil,ParamSoc,Hstatus,
  UtilXlsBTP,EtudesStruct,EtudePiece,FactComm, ComCtrls,
  //uMetreArticle,
  uEntCommun
  ;
type

cleOle = record
  NumOle : integer;
end;

TStockMetre = class
private
	//RepertTmp   : String;

  //Creation de la TOB des Variables Application/Document/Article
  TOBVariables: TOB;
  //
  fUSable :Boolean;
  fRepertBase : String;

  // Modif metre => FV
  SNaturePiece : string;
  SSouche : string;
  SNumeroPiece : integer;
  SInd : Integer;

  LeFichier : string;
  Rep_Ouvrage : string;
  fNomFic : string; // Nom du fichier
  fRepert : string; // repertoire de stockage pour le doc ou l'article (nom du sous repert)
  fuser : string; // Utilisateur (Obligatoire);
  fRepertLocal : String; //Repertoire Local pour sauvegarde fichier
  FNomMachine : string; //Nom de la machine utilisateur (Obligatoire)
  fAction : TActionFiche;
  fcreat : boolean;
  fDuplication : boolean;
  ThisLigneDoc : TOB; // Ligne Courante Document
  ThisLigneOLE : TOB; // Ligne Courante OLE
  ThisOuvrages : TOB; // les Ouvrages
  ThisDoc : TOB;   // Document Courant
  sTOBMetres,sTobMetres_O : TOB; // TOB des metres
  IndiceOle : INTEGER;
  fWinExcel: OleVariant;
  AValider : boolean;
  //
  function  findLaLigneInit (ThisTOB : TOB) : TOB;
  function  findLaLigneOLE(ThisTOB: TOB): TOB;
  procedure PositionneLeDoc (DocAFind : TOB);
  procedure PositionneLaLigne (LigneAFind : TOB);
  procedure CreeLienOle;
  function  GenereNomFichier (ThisTOB : TOB) : string;
  procedure AttachesOleOuvrages (TOBOuvrages : TOB; var II : integer);
  function  decodeLeNomFichier(TheNomFic: string): cleOle;
  function GenereLaRefMetre(ThisTOB: TOB): string;
  procedure EpurationRepertoire;
  function FichierTrouve (TheFichier : string) : boolean;
  procedure addlessupOle (TheOle : TOB);
  procedure Desactive (laForm : TForm);
  procedure Active (laForm : TForm);
  procedure NettoieSauvegarde;
  procedure RestitueSauvegarde;
  procedure EpureRepertoireOuvrage;
  // Gestion des Variables
  function ControleText(Libelle : string): Boolean;

  function ExtraireTexte(Source, TexteDep, Limites: string; Depart, Deb, Longueur: integer): string;
  function RecupVariableInFic (Nom_Fic_Text:string) : Variant;
  Function RecupValeurMetre(TOBL : TOB) : Variant;
  Function OkSaisieVariables() : Boolean;
  procedure AddChampsSupVar (TOBFille : TOB);
{$IFDEF CEAGLCLIENT}
  procedure SetTOBMetres (TOBmetre : TOB);
{$ENDIF}

public

  TOBSaisieMetre: TOB;

  constructor create ;
  destructor destroy ; override;
  property User : string read fuser write fuser;
  property Ligne : TOB read ThisLigneOLE write PositionneLaLigne ;
  property Document : TOB read ThisDoc write PositionneLeDoc;
  property Ouvrages : TOB write ThisOuvrages;
  property Fichier : string read fNomfic;
  property Usable : boolean read fUsable;
  property Action : TActionFiche read Faction write Faction;
  property Duplic : boolean read Fduplication write Fduplication;
{$IFDEF CEAGLCLIENT}
  Property TOBMetre : TOB read sTOBMetres write SetTOBMetres;
{$ELSE}
  Property TOBMetre : TOB read sTOBMetres;
{$ENDIF}


  //
  function ModifieMetre (TOBL : TOB; TheForm : TOBject;CreationAutorise : boolean=true) : double;
  function CalculeMetre (TOBL : TOB; TheForm : TOBject) : double;
  function GetValeurMetre (TOBL : TOB;Creation : boolean=true) : variant;
  function QteModifie (TOBL : TOB) : boolean;
  function MetreAttached (TOBL : TOB) : boolean;
  procedure SauvegardeFichiers;
  procedure Validate;
  procedure Destruct;
  procedure Annulation;
  procedure ControleRepertoire;
  procedure RecupMetreBiblio(TOBL:TOB; FicArt: string);
  procedure SaisieVar(TOBL: TOB);
  procedure LanceVar(Action : TactionFiche);
  Procedure CopieSousDetail();
  function MetreSousDetail(TOBL: TOB): boolean;
  procedure DocMetreBiblio(TOBL:TOB);
  // Gestion des Variables
  procedure EcrireText();
  procedure CreationVarDoc(CodeArt,NaturePiece, Souche : string; IO, NumeroPiece, I : Integer);
  procedure ChargementVariables(Cledoc : R_CleDoc);
  Function ControlMetre: boolean;
  Function AutorisationMetre(NatPiece : string) : Boolean;
end;

var TheMetreShare : TStockMetre;

implementation

uses facture,Lignomen,factUtil,UmetresUtil;

function TStockMetre.GenereLaRefMetre(ThisTOB : TOB): string;
begin
	result := GenereLaReferenceMetre (ThisTOB,Action,Avalider);
end;


{ TMetre }

function TStockMetre.decodeLeNomFichier (TheNomFic : string) : cleOle;
var SLocal: string;
    CleLoc : CleOle;
    leNomFic : string;
begin
  FillChar(Cleloc,Sizeof(CleLoc),#0) ;
  Result := CleLoc;
  LeNomFic := ExtractFileName(TheNomfic);
  Slocal := copy (LeNomFic,1,pos('.xls',LeNomfic)-1);

  if Slocal = '' then exit;

  if sLocal <> '' then CleLoc.NumOle := strToInt(Slocal);

  Result := CleLoc;

end;

function TStockMetre.GenereNomFichier(ThisTOB: TOB): string;
begin
  Result := inttostr(ThisTOb.GetValue('INDICEMETRE'))+'.xls';
end;

constructor TStockMetre.create;
begin

  IndiceOle := 0;
  fusable := false;
  fDuplication := false;
  ThisOuvrages := nil;
  fWinExcel := Unassigned;
  fcreat := true;
  AValider := false;
  //
  sTOBMetres := TOB.Create('LES METRES',nil,-1);
  sTOBMetres_O := TOB.Create('LES METRES ORIGINES',nil,-1);
  TobSaisieMetre := TOB.Create('Var ResultMetre', nil, -1);

  FNomMachine := ComputerName ;

  if GetParamSoc ('SO_BTMETREDOC') then frepertBase := RecupRepertbase
  else frepertbase := '';

  if frepertBase = '' Then
  begin
    exit;
  end;

  NettoieSauvegarde;

  fUsable := true;

end;

procedure TStockMetre.CreeLienOle;
var LARef : string;
begin
  Inc(IndiceOle); // c'est une creation donc on incremente le IndiceOle
  ThisLigneDoc.PutValue('INDICEMETRE',IndiceOle); // Ligne de document
  LaRef := GenereLaRefMetre(ThisLigneDoc);
  LeFichier := GenereNomFichier (ThisLigneDoc);

  //Creation du répertoire Nouveau-0-0
  fNomFic := fRepertBase+'\'+fRepert+'\';

  if not DirectoryExists(FNomFic) then
     begin
     if not Creationdir(FNomFic) then
        BEGIN
        PGIBoxAF(traduirememoire('Répertoire de travail inexistant : ') + FNomfic + chr(10),traduirememoire('METRE'));
        exit;
        END;
     end;
//
  fNomFic := fRepertBase+'\'+fRepert+'\'+LeFichier;
//
  ThisLigneOle := TOB.Create('LIENSOLE',sTOBMetres,-1);
  addlessupOle (ThisLigneOle);
  ThisLigneOle.putvalue('LO_TABLEBLOB','MTR');
  ThisLigneOle.putvalue('LO_QUALIFIANTBLOB','TXT');
  ThisLigneOle.putvalue('LO_EMPLOIBLOB','DIV');
  ThisLigneOle.putvalue('LO_RANGBLOB',1);
  ThisLigneOle.putvalue('LO_LIBELLE','Métrés');
  ThisLigneOle.putvalue('LO_IDENTIFIANT',LaRef);
  ThisLigneOle.putvalue('LO_OBJET',fNomFic);
  ThisLigneOle.putvalue('LO_PRIVE','-');
  ThisLigneOle.putvalue('LO_DATEBLOB',V_PGI.DateEntree);
  ThisLigneOle.putvalue('LO_CREATEUR',V_PGI.User);
  ThisLigneOle.putvalue('LO_UTILISATEUR',V_PGI.User);
  ThisLigneOle.putvalue('INDICEMETRE',IndiceOle);
  //
end;

destructor TStockMetre.destroy;
begin
  if ((Action = taCreat) and (not AValider )) then deleteRepertoire ( fRepertBase+'\'+fRepert);
  sTOBMetres.free;
  sTOBMetres_o.free;
  TobSaisieMetre.free;
  NettoieSauvegarde;
  inherited;
end;

function TStockMetre.findLaLigneInit(ThisTOB: TOB): TOB;
var LaRef : string;
begin
  LaRef := GenereLaRefMetre (ThisTOB);
  result := sTOBMetres.findfirst (['LO_IDENTIFIANT'],[LaRef],false);
end;

function TStockMetre.findLaLigneOLE(ThisTOB: TOB): TOB;
begin

  if ThisTOB.getValue('INDICEMETRE') = 0 then
     BEGIN
       Result := nil ;
       Exit;
     END;
  result := sTOBMetres.findfirst (['INDICEMETRE'],[ThisTOB.getValue('INDICEMETRE')],true);

end;   

procedure TStockMetre.PositionneLaLigne (LigneAFind : TOB);
begin
  ThisLigneOle := nil;
  if not Usable then exit;
  if LigneAfind = nil then exit;
  if (LigneAFind.NomTable  <> 'LIGNE') and (LigneAFind.NomTable <> 'LIGNEOUV') then exit;
  ThisLigneDoc := LigneAFind; // ligne de document ou d'ouvrage
  ThisLigneOle := findLaLigneOLE (ThisLigneDoc);
  if (ThisLigneOle = nil) and (fCreat) Then CreeLienOle;
  if thisLigneOle <> nil then
     fNomFic := Trim(ThisLigneOLE.GetValue('LO_OBJET'))
  else
     fNomFic := '';
end;

procedure TStockMetre.PositionneLeDoc(DocAFind: TOB);
var LaRef : string;
    Select : string;
    QQ : TQuery;
    Indice,II : integer;
    TOBL,TOBOLE : TOB;
    r_cleOle : cleOle;
begin

  if not Usable then exit;

  if DocAfind = nil then exit;

  if fUser = '' then
     BEGIN
     fusable := false;
     exit;
     END;

  if (DocAfind.NomTable  <> 'PIECE') and
     (DocAfind.NomTable <> 'ARTICLE') and
     (DocAFind.nomTable <> 'NOMENENT') then
     BEGIN
     fUsable := false;
     exit;
     END;

  ThisDoc := DocAfind; // pointe sur la Piece ou l'article ou la nomenclature
  LaRef := GenereLaRefMetre (DocAFind);
  fRepert := LaRef;
  fRepertLocal := FNomMachine + '-' + fUser;
  SauvegardeFichiers;
  //
  IndiceOle := 0;
  II := 0;
  sTOBMetres_O.ClearDetail;
  sTOBMetres.ClearDetail;
  //
  Select := 'SELECT * FROM LIENSOLE WHERE (LO_TABLEBLOB="MTR") AND (LO_IDENTIFIANT LIKE "'+LaRef + '%")';
  QQ := OpenSql (Select,true);
  if not QQ.eof then sTOBMetres.LoadDetailDB ('LIENSOLE','','',QQ,false);
  ferme (QQ);
  sTOBMetres_O.Dupliquer (sTOBmetres,true,true);

  // Ajout des champs sup obligatoires
  for Indice := 0 to sTOBMetres.detail.count -1 do
      begin
      TOBOLE := sTOBMetres.detail[Indice];
      addlessupOle (TOBOLE);
      end;

  // Attache les lignes au lien OLE correspondant
  For Indice := 0 to ThisDoc.detail.count -1 do
      begin
      TOBL := ThisDoc.detail[Indice];
      TOBOLE := findLaLigneInit (TOBL);
      if TOBOLE <> nil then
         begin
           Inc( II);
//      r_cleOle := decodeLeNomFichier (Trim(TOBOLE.GetValue('LO_OBJET')));
         TOBL.PutValue('INDICEMETRE',II); // Ligne de document
         TOBOLE.PutValue('INDICEMETRE',II); // LIENSOLE
         if II > IndiceOle then IndiceOle := II;
         end;
      end;

  if ThisOuvrages <> nil then
     begin
    AttachesOleOuvrages (ThisOuvrages,II);
     end;

end;

procedure TStockMetre.AttachesOleOuvrages (TOBOuvrages : TOB; var II : integer);
var Indice : integer;
    TOBL, TOBOLE : TOB;
    r_CleOle : cleOle;
begin

  For Indice := 0 to TOBOuvrages.detail.count -1 do
      begin
      TOBL := TOBOuvrages.detail[Indice];
    if TOBL.detail.count > 0 then AttachesOleOuvrages (TOBL,II);
      if (TOBL.NomTable <> 'LIGNEOUV') and (TOBL.NomTable <> 'NOMENLIG') Then Continue;
      TOBOLE := findLaLigneInit (TOBL);
      if TOBOLE <> nil then
         begin
      Inc( II);
//      r_cleOle := decodeLeNomFichier (Trim(TOBOLE.GetValue('LO_OBJET')));
      TOBL.PutValue('INDICEMETRE',II); // Ligne de document
      TOBOLE.PutValue('INDICEMETRE',II); // LIENSOLE
      if II > IndiceOle then IndiceOle := II;
         end;
      end;

end;

function TStockMetre.ModifieMetre (TOBL : TOB; TheForm : TOBject;CreationAutorise : boolean) : double;
var laForm      : TForm;
    RepertReel  : string;
    SavFcreat   : boolean;
begin

  result := 0;

  SavFcreat := fcreat;    // Mise de cote pour restauration Ulterieure

  fcreat := CreationAutorise;

  if fUser = '' then
     Begin
     fusable := false;
     exit;
     End;

  if not Usable then exit;

  //
  RepertReel := fRepertBase+'\'+fRepert;

  if not DirectoryExists(RepertReel) then
     begin
     if not Creationdir(repertReel) then
        Begin
        fUsable := false;
        exit;
        End;
     end;

  laForm := TForm(TheForm); // TFFActure ou TFLIgNomen
  Desactive (laForm);

  TRY
    PositionneLaLigne (TOBL);
    if (ThisLigneOLE = nil) or (fNomFic = '') then exit;
    if not OfficeExcelDispo then
       begin
       PGIBoxAF (TraduireMemoire('Office n''est pas installé sur ce poste'),traduireMemoire('METRE'));
       exit;
       end;
    if fNomFic='' then exit;
    if not FileExists (fNomFic) then
       Begin
       // Création du fichier Variables.txt en fonction de BVARDOC du document
       SaisieVar(TOBL);
      if not CreateTheFichierXls (fnomfic) then exit;
       end;
    AddVbModuleFromFile ('c:\PGI00\STD\', 'MetrePgi.xla');
    FileExecAndWait (GetExcelPath + 'EXCEL.exe "'+fNomFic+'"');
    Result := GetValeurMetre (TOBL);
    if FileExists ('C:\pgi00\STD\Variables.txt') then
       Begin
       // suppression du fichier Variables
         DeleteFichier('C:\pgi00\STD\Variables.txt');
       End;

    // suppression du fichier macro
    DelVBModule ('MetrePgi.xla');

  FINALLY
    Active (laForm);
    fcreat := SavFcreat;
  END;

end;

function TStockMetre.GetValeurMetre (TOBL : TOB;Creation : boolean=true) : variant;
var WinNew : boolean;
    WorkBook,Mns : Variant;
    Indice : integer;
    SavFcreat : boolean;
    Nom_Fic_Text : string;
begin

  Result := 0;
  SavFcreat := false;

  if not USable then Exit;

  if not creation then
     begin
     SavFcreat := fcreat;
     fcreat := creation;
     end;

  PositionneLaLigne (TOBL);

  if not creation then
    begin
      Fcreat := SavFcreat;
    end
  else
    begin
      if not FileExists (fNomFic) then
        If Not CreateTheFichierXls (fnomfic) then exit;
    end;

  if fnomfic = '' then exit;

  Nom_Fic_Text := copy (fnomfic,1,Pos('.',FnomFic)-1)+'.txt';

  // controle si fichier txt ou XLS existe
  if FileExists (Nom_fic_text) then
     begin
     result := RecupVariableInFic (Nom_Fic_Text);
     DeleteFichier (Nom_Fic_Text);
     DelVBModule ('MetrePgi.xla');
     Exit;
     end
  Else
    Begin
     if not fileExists (fnomfic) then Exit;
    end;

  // A completer recup de la valeur

(* ce n'est plus utile de tester ceci
  if IsExcelLaunched then
     BEGIN
       PGIBox (TraduireMemoire('Vous devez fermer préalablement les instances d''EXCEL'), TraduireMemoire('METRE'));
     END;
*)
  if not OpenExcel (true,fWinExcel,WinNew) then
     begin
     PGIBox (TraduireMemoire('Liaison Excel impossible'), TraduireMemoire('Liaison EXCEL'));
     fNomFic := '';
     DelVBModule ('MetrePgi.xla');
     exit;
     end;

  TRY
    WorkBook := OpenWorkBook (FNomFic ,fWinExcel);
    Mns := WorkBook.names;
    if Mns.count > 0 then
       begin
       for Indice := 1 to mns.count do
           begin
           if WorkBook.names.item(indice).name = 'METRERESULTAT' then
              begin
              if IsNumeric(WorkBook.names.item(Indice).refersToRange.text) then
                 Result := WorkBook.names.item(Indice).refersToRange.Value
              else
                 Result := 0;
              Break;
              end;
           end;
       end;
  FINALLY
    ExcelClose (fwinExcel);
  END;

  DelVBModule ('MetrePgi.xla');

//
end;

procedure TStockMetre.Validate;
var lastrep,Fichier,LeNewFic,LeOldFic,st : string;
    LaTobMere,LaTOBOLe,TOBL : TOB;
    Indice : integer;
begin

  if fUser = '' then
     BEGIN
     fusable := false;
     exit;
     END;

  if not usable then exit;

  if (sTobMetres_O.detail.count > 0) and (not Duplic) then
     begin
     if not (ExecuteSQL ('DELETE FROM LIENSOLE WHERE (LO_TABLEBLOB="MTR") AND (LO_IDENTIFIANT LIKE "'+ fRepert + '%")')>=0) then
        BEGIN
        MessageValid := 'Erreur Supression METRES';
        V_PGI.IOError := oeUnknown;
        Exit;
        END;
     end;

  if TobSaisieMetre = nil then exit;

  St := ' BVD_NATUREPIECE="' + ThisDoc.GetValue('GP_NATUREPIECEG') + '"'+
        ' AND BVD_SOUCHE="' + ThisDoc.GetValue('GP_SOUCHE') + '" ' +
        ' AND BVD_NUMERO=' + InttoStr(ThisDoc.GetValue('GP_NUMERO')) +
        ' AND BVD_INDICE=' + IntToStr(ThisDoc.GetValue('GP_INDICEG')) + ' ';
  if ExecuteSql('DELETE FROM BVARDOC WHERE '+st) < 0 then
     begin
     MessageValid := 'Erreur Supression VARIABLES DOC';
     V_PGI.IOError := oeUnknown;
     Exit;
     end;

  //Chargement des enregistrements de TOBSaisieMetre vers Table BVARDOC
  for Indice := 0 to TOBSaisieMetre.detail.count -1 do
      begin
      TOBL := TOBSaisieMetre.detail[Indice];
      TOBL.Putvalue('BVD_NATUREPIECE',ThisDoc.GetValue('GP_NATUREPIECEG'));
      TOBL.Putvalue('BVD_SOUCHE',ThisDoc.GetValue('GP_SOUCHE'));
      TOBL.Putvalue('BVD_NUMERO',ThisDoc.GetValue('GP_NUMERO'));
      TOBL.Putvalue('BVD_INDICE',ThisDoc.GetValue('GP_INDICEG'));
      TOBL.setallModifie(true);
      end;

  if not TobSaisieMetre.InsertDB(nil,true) then
     begin
     MessageValid := 'Erreur écriture VARIABLES';
     V_PGI.IOError := oeUnknown;
     Exit;
     end;

  if (Action = taCreat) and (sTObMetres.detail.count >0) then
     begin
     LastRep := frepert;
     AValider := true;
     fRepert := GenereLaRefMetre (ThisDoc); // defini le repert reel de stockage
    if not renameFichier (fRepertBase+'\'+LastRep,fRepertBase+'\'+fRepert) then
        BEGIN
        fRepert := LastRep;
        AValider := false;
        MessageValid := 'Erreur écriture fichiers métrés';
        V_PGI.IOError := oeUnknown;
        Exit;
        END;
     end;

  //Cas de la duplication on cree le repertoire destination
  if (Duplic) and (sTObMetres.detail.count >0) Then
     begin
     fRepert := GenereLaRefMetre (ThisDoc); // defini le repert reel de stockage
    if not DirectoryExists (fRepertBase+'\'+frepert ) then
        begin
      Creationdir (fRepertBase+'\'+frepert );
        end;
     end;

  // PArcours de la TOB sTOBMetres pour eliminer les fichiers en trop et remettre a niveau la tOB
  Indice := 0;
  if sTobMetres.detail.count > 0 then
     begin
     repeat
     LaTOBOle := sTOBMetres.detail[Indice];
     // recherche sur les lignes de documents (LIGNES)
     laTOBMere := ThisDoc.findFirst(['INDICEMETRE'],[laTOBOle.GetValue('INDICEMETRE')],true);
     if laTOBMere <> nil then
        begin
        laTOBOle.putvalue('LO_IDENTIFIANT',GenereLaRefMetre(laTOBMere));
        if Action = taCreat then
           begin
           leOldFic := LaTOBOle.GetValue('LO_OBJET');
           LenewFic := GenereNomFichier(LaTOBMere);
           //FV1 : 29/01/2014 - FS#844 - ROUSSEL PEINTURE - message "E/S 103" en duplication de pièces
         if FileExists(LeOldFic) then renameFichier (leOldFic,fRepertBase+'\'+fRepert+'\'+LeNewFic);
         //FV1 -  05/01/2015 - FS#1247 - VEODIS GROUP - métrés
             laTOBOle.putvalue('LO_OBJET',fRepertBase+'\'+fRepert+'\'+LeNewFic);
           end;
       //Duplication
        if Duplic then
           begin
           leOldFic := LaTOBOle.GetValue('LO_OBJET');
           LenewFic := GenereNomFichier(LaTOBMere);
           //FV1 : 29/01/2014 - FS#844 - ROUSSEL PEINTURE - message "E/S 103" en duplication de pièces
         if FileExists(LeOldFic) then CopieFichier (leOldFic,fRepertBase+'\'+fRepert+'\'+LeNewFic);
         //FV1 -  05/01/2015 - FS#1247 - VEODIS GROUP - métrés
             laTOBOle.putvalue('LO_OBJET',fRepertBase+'\'+fRepert+'\'+LeNewFic);
           end;
        laTOBOLe.setallModifie(true);
        Inc(Indice);
     end
     else
        begin
        // recherche sur les lignes d'ouvrages
        laTOBMere := ThisOuvrages.findFirst(['INDICEMETRE'],[laTOBOle.GetValue('INDICEMETRE')],true);
        if laTOBMere <> nil then
           begin
           laTOBOle.putvalue('LO_IDENTIFIANT',GenereLaRefMetre(laTOBMere));
           if Action = taCreat then
              begin
              leOldFic := LaTOBOle.GetValue('LO_OBJET');
              LenewFic := GenereNomFichier(LaTOBMere);
            //FV1 : 29/01/2014 - FS#844 - ROUSSEL PEINTURE - message "E/S 103" en duplication de pièces
           if FileExists(LeOldFic) then renameFichier (leOldFic,fRepertBase+'\'+fRepert+'\'+LeNewFic);
           //FV1 -  05/01/2015 - FS#1247 - VEODIS GROUP - métrés
               laTOBOle.putvalue('LO_OBJET',fRepertBase+'\'+fRepert+'\'+LeNewFic);
             end;
         //
           if Duplic then
              begin
              leOldFic := LaTOBOle.GetValue('LO_OBJET');
              LenewFic := GenereNomFichier(LaTOBMere);
              //FV1 : 29/01/2014 - FS#844 - ROUSSEL PEINTURE - message "E/S 103" en duplication de pièces
           if FileExists(LeOldFic) then CopieFichier (leOldFic,fRepertBase+'\'+fRepert+'\'+LeNewFic);
           //FV1 -  05/01/2015 - FS#1247 - VEODIS GROUP - métrés
                laTOBOle.putvalue('LO_OBJET',fRepertBase+'\'+fRepert+'\'+LeNewFic);
              end;
           laTOBOLe.setallModifie(true);
           Inc(Indice);
       end
       else
        begin
           Fichier := Trim(laTOBOle.GetValue('LO_OBJET'));
           if FileExists (Fichier) then DeleteFichier (Fichier);
           laTOBOle.free;
           end;
        end;
     until Indice >= sTOBMetres.detail.count;
     if not sTOBMetres.InsertDB (nil,true) then
        BEGIN
        MessageValid := 'Erreur écriture INFOS METRES';
        V_PGI.IOError := oeUnknown; Exit;
        END;
     end;

  SauvegardeFichiers;

  if Action = taModif then EpurationRepertoire;
  sTobMetres_O.clearDetail;
  sTOBMetres_O.Dupliquer(sTOBMetres,true,true); 

end;

procedure TStockMetre.EpurationRepertoire;
var NomFic : string;
    Rec : TSearchRec;
begin
  Nomfic:= fRepertBase+'\'+fRepert+'\*.*';
  if FindFirst (Nomfic,faAnyFile,Rec) = 0 then
     begin
     if (rec.name <> '.') and (rec.name <> '..') then
      if not FichierTrouve (rec.name) then deleteFichier (fRepertBase+'\'+fRepert+'\'+rec.Name);
     while FindNext (Rec) = 0 do
        begin
        if (rec.name <> '.') and (rec.name <> '..') then
        if not FichierTrouve (rec.name) then deleteFichier (fRepertBase+'\'+fRepert+'\'+rec.Name);
        end;
     end;

  FindClose (Rec);

end;

function TStockMetre.FichierTrouve(TheFichier: string): boolean;
var laTOB : TOB;
    
begin
  laTOB := sTOBMetres.findFirst(['LO_OBJET'],[fRepertBase+'\'+fRepert+'\'+TheFichier],true);
  result := (LaTob <> nil);
end;

procedure TStockMetre.Destruct;
Var St    : String;
begin

  if sTobMetres_O.detail.count > 0 then
     begin
     if not (ExecuteSQL ('DELETE FROM LIENSOLE WHERE (LO_TABLEBLOB="MTR") AND (LO_IDENTIFIANT LIKE "'+ fRepert + '%")')>=0) then
        BEGIN
        V_PGI.IOError := oeUnknown;
        Exit;
        END;
     if TobSaisieMetre <> nil then
        Begin
        St := ' BVD_NATUREPIECE="' + ThisDoc.GetValue('GP_NATUREPIECEG') + '"'+
              ' AND BVD_SOUCHE="' + ThisDoc.GetValue('GP_SOUCHE') + '" ' +
              ' AND BVD_NUMERO=' + InttoStr(ThisDoc.GetValue('GP_NUMERO')) +
              ' AND BVD_INDICE=' + IntToStr(ThisDoc.GetValue('GP_INDICEG')) + ' ';
        if ExecuteSql('DELETE FROM BVARDOC WHERE '+st) < 0 then
           begin
           V_PGI.IOError := oeUnknown;
           Exit;
           end;
        TobSaisieMetre.clearDetail;
        end;
     end;

  sTobMetres.clearDetail;
  sTobMetres_O.clearDetail;
  deleteRepertoire ( fRepertBase+'\'+fRepert);
end;

procedure TStockMetre.addlessupOle(TheOle: TOB);
begin
  TheOle.addchampsupValeur('INDICEMETRE',0,false);
end;

procedure TStockMetre.Annulation;
begin

  if fUser = '' then
     BEGIN
     fusable := false;
     exit;
     END;

  if not Usable then exit;

  sTobMetres.cleardetail ;

  sTobMetres.dupliquer(sTobMetres_O,true,true);

  RestitueSauvegarde;

  ControleRepertoire;

end;

procedure TStockMetre.ControleRepertoire;
var NomFic   : string;
    Rec : TSearchRec;
begin

  Nomfic := fRepertBase + '\' + fRepert + '\*.*';

  if FindFirst (Nomfic,faAnyFile,Rec) = 0 then
     begin
     if (rec.name <> '.') and (rec.name <> '..') then
        begin
        FindClose (Rec);
        exit;
        end;
     while FindNext (Rec) = 0 do
        begin
        if (rec.name <> '.') and (rec.name <> '..') then
           begin
           FindClose (Rec);
           exit;
           end;
        end;
     end;

  FindClose (Rec);

  // Suppression repertoire Ouvrage
  deleteRepertoire(fRepertBase + '\' + fRepert + '\');

end;

procedure TStockMetre.Active(laForm: TForm);
begin

  if (laForm.ClassName) = 'TFFacture' then
     begin
     if TFFacture(laForm).Pentete.visible then TFFacture(laForm).PEntete.Enabled := true;
     if TFFacture(laForm).PenteteAffaire.visible then TFFacture(laForm).PEnteteAffaire.Enabled := true;
//   if TFFacture(laForm).Debut.visible then TFFacture(laForm).Debut.Enabled := true;
//   if TFFacture(laForm).Fin.visible then TFFacture(laForm).Fin.Enabled := true;
     TFFacture(laForm).PPied.Enabled := true;
     TFFacture(laForm).DockBottom.Enabled := true;
     TFFacture(laForm).GS.Enabled := true;
     TFFActure(laform).gs.SetFocus;
     end
  else if (laForm.classname) = 'TFLigNomen' then
     begin
     if TFLigNomen(laForm).TV_NLIG.visible then TFLigNomen(laForm).TV_NLIG.enabled := true;
     TFLigNomen(laForm).Dock971.Enabled := true;
     TFLigNomen(laForm).G_NLIG.Enabled := true;
     TFLigNomen(laform).G_NLIG.SetFocus;
     end
  else if UpperCase(laForm.className) = 'TFNOMENLIG' then
  // Detail d'ouvrage dans bibliotheque
     begin
     end;

end;

procedure TStockMetre.Desactive(laForm: TForm);
begin

  if laForm.ClassName = 'TFFacture' then
     begin
     if TFFacture(laForm).Pentete.visible then TFFacture(laForm).PEntete.Enabled := false;
     if TFFacture(laForm).PenteteAffaire.visible then TFFacture(laForm).PEnteteAffaire.Enabled := false;
//   if TFFacture(laForm).Debut.visible then TFFacture(laForm).Debut.Enabled := false;
//   if TFFacture(laForm).Fin.visible then TFFacture(laForm).Fin.Enabled := false;
     TFFacture(laForm).PPied.Enabled := false;
     TFFacture(laForm).DockBottom.Enabled := false;
     TFFacture(laForm).GS.Enabled := false;
     end
  else if laForm.classname = 'TFLigNomen' then
     begin
     TFLigNomen(laForm).Dock971.Enabled := false;
     TFLigNomen(laForm).G_NLIG.Enabled := false;
     if TFLigNomen(laForm).TV_NLIG.visible  then TFLigNomen(laForm).TV_NLIG.enabled := false;
     end;

end;

function TStockMetre.QteModifie(TOBL: TOB): boolean;
begin

  if fUser = '' then
     BEGIN
     fusable := false;
     result  := false;
     exit;
     END;

  result := true;
  fcreat := false;
  PositionneLaLigne (TOBL); // juste positionnement
  fcreat := true;

  if (ThisLigneOLE = nil) or (fNomFic = '') then exit;

  if PGIAsk (traduireMemoire('Confirmez-vous la suppression du métré ?'),TraduireMemoire('METRE')) <> mrYes then
     begin
     result := false;
     exit;
     end;

  TOBL.putValue('INDICEMETRE',0);

end;

procedure TStockMetre.NettoieSauvegarde;
begin
  if fUser = '' then exit;
  NettoieSauv (fRepertBase,'Sauve',fRepertLocal);
end;

procedure TStockMetre.SauvegardeFichiers;
begin
  SauvegardeFic (frepertBase,frepert,'Sauve',fRepertLocal);
end;

procedure TStockMetre.RestitueSauvegarde;
begin
  RestitueSauve (frepertBase,frepert,'Sauve',fRepertLocal);
end;

function TStockMetre.MetreAttached(TOBL: TOB): boolean;
begin

  if fUser = '' then
     BEGIN
     fusable := false;
     result  := false;
     exit;
     END;

  result := false;

  if not Usable then exit;

  ThisLigneOle := findLaLigneOLE (TOBL);

  if (ThisLigneOLE <> nil) then Result := True;

end;

// Lancement saisie des variables et ecriture fichier Txt
procedure TStockMetre.SaisieVar(TOBL: TOB);
Var Prefixe: string;
begin

  if TOBL.NomTable = 'LIGNE' then
    Prefixe := 'GL_'
  Else if TOBL.NomTable = 'LIGNEOUV' Then
    Prefixe := 'BLO_'
  Else
    Exit;

  // Création du fichier Variables.txt en fonction de BVARDOC du document
  EcrireText();

end;

// Copie du fichier metre de la bibliothèque en fichier metre de la ligne documents
procedure TStockMetre.RecupMetreBiblio(TOBL: TOB; FicArt: string);
Var CodeArt     : String;
    TypeArt     : String;
    Prefixe     : string;
    NaturePiece : String;
    Souche : String;
    NumeroPiece : Integer;
    Indice : Integer;
    NumOrdre    : Integer;
    UniqueBlo   : Integer;
    //IndiceFile : Integer;
begin

  if TOBL.NomTable = 'LIGNE' then
    Prefixe := 'GL_'
  Else if TOBL.NomTable = 'LIGNEOUV' Then
    Prefixe := 'BLO_'
  Else
    Exit;

  TypeArt := TOBL.GetValue(Prefixe + 'TYPEARTICLE');
  CodeArt := TOBL.GetValue(Prefixe + 'ARTICLE');

  //Chargement de la cle servant à identifier le document
  NaturePiece := TOBL.GetValue(Prefixe + 'NATUREPIECEG');
  Souche := TOBL.GetValue(Prefixe + 'SOUCHE');
  NumeroPiece := TOBL.GetValue(Prefixe + 'NUMERO');
  Indice := TOBL.GetValue(Prefixe + 'INDICEG');
  NumOrdre    := TOBL.GetValue(Prefixe + 'NUMORDRE');
  if TOBL.NomTable = 'LIGNEOUV' Then   UniqueBlo    := TOBL.GetValue(Prefixe + 'UNIQUEBLO');

  //Copie du fichier de départ article en fichier index métrés
  if FicArt <> '' then
     begin
     // Recherche du N° d'index du fichier métrés
     PositionneLaLigne(TOBL);
     CopieFichier(FicArt,FNomFic);
     end;

  //Appel de la procedure de copie des variables d'un article
  // dans les variables d'un document
  if GetParamSoc('SO_BTMETREDOC') then
	   CreationVarDoc(CodeArt,NaturePiece, Souche, IndiceOle, NumeroPiece,Indice);

end;

// Copie des fichiers des sous détail dans un répertoire ouvrage
procedure TStockMetre.CopieSousDetail();
begin

  // chargement du répertoire spécial pour le stockage des XLS du sous détail
  rep_ouvrage := frepertBase + '\Ouvrage';

  if not DirectoryExists (Rep_Ouvrage) then
     begin
     Creationdir (Rep_Ouvrage);
     end;

  CopieFichier(fNomfic, rep_ouvrage + '\' + LeFichier);

end;

function TStockMetre.CalculeMetre(TOBL: TOB; TheForm: TOBject): double;
var RepertReel  : string;
//    RepertTmp   : String;
    SavFcreat   : boolean;
    Prefixe     : string;
begin

  result := 0;
  Prefixe := '';
  SavFcreat := fcreat;    // Mise de cote pour restauration Ulterieure

  if fUser = '' then BEGIN fusable := false; exit; END;

  if not Usable then exit;
  //
  RepertReel := fRepertBase+'\'+fRepert;

  if not DirectoryExists (RepertReel) then
     begin
     if not Creationdir (repertReel) then BEGIN fUsable := false;exit; END;
     end;

  //laForm := TForm(TheForm); // TFFActure ou TFLIgNomen
  TRY
    fcreat := false;
    PositionneLaLigne (TOBL);
    if (ThisLigneOLE = nil) or (fNomFic = '') then exit;
    if TOBL.NomTable = 'LIGNE' then
       Prefixe := 'GL_'
    Else if TOBL.NomTable = 'LIGNEOUV' Then
       Prefixe := 'BLO_'
    Else
       Exit;
    if TOBL.GetValue(Prefixe + 'TYPEARTICLE') = 'OUV' then
       Begin
       Result := RecupValeurMetre(TOBL);
       end
    Else
       Begin
       if not OfficeExcelDispo then
          begin
          PGIBoxAF (traduirememoire('Office n''est pas installé sur ce poste'),traduireMemoire('METRE'));
          exit;
          end;
       if fNomFic='' then exit;
       if not FileExists (fNomFic) then
       begin
         if not CreateTheFichierXls (fnomfic) then exit;
       end;
       AddVbModuleFromFile ('c:\PGI00\STD\','MetrePgi.xla');
       FileExecAndWait (GetExcelPath + 'EXCEL.exe "' + fNomFic +'"');
       Result := RecupValeurMetre(TOBL);
       End;
  FINALLY
  fcreat := SavFcreat;
  END;

end;

// fonction permettant la récupération de la valeur dans fichier Txt
// ou  bien dans fichier XLS.
Function TStockMetre.RecupValeurMetre(TOBL : TOB) : Variant;
Begin

  result := GetValeurMetre (TOBL);

  TRY
    result := Arrondi(VarasType(result,vardouble),V_PGI.OkDecQ);
  EXCEPT
    on EVariantError do PGIBox (TraduireMemoire('La valeur n''est pas numérique'),TraduireMemoire('METRES'));
  END;

  // Rajout LS
  // Modification FV pour annulation suppression du fichier XLS
  //if Result = 0 then TOBL.PutValue('INDICEMETRE',0);
  // --
  // suppression du fichier des variables
  //if FileExists (fRepertBase+ '\'+fRepert + '\Variables.txt') then


  if FileExists ('C:\pgi00\STD\Variables.txt') then
     Begin
       // suppression du fichier Variables
       DeleteFichier('C:\pgi00\STD\Variables.txt');
     End;

  // suppression du fichier macro
  DelVBModule ('MetrePgi.xla');

End;


// permet de lancer EXCELL en automatique dans le cas d'un sous-détail d'ouvrage
// avec des métrés
function TStockMetre.MetreSousDetail(TOBL: TOB): boolean;
var Fic_Bidon : string;
Begin
  result := false;
  if not Usable then exit;

  // Verification si le répertoire "C:\.. ..\Ouvrage existe
  // S'il existe pas on sort
  if not DirectoryExists (Rep_Ouvrage) then exit;

  // TFFActure ou TFLIgNomen
  //laForm := TForm(TheForm);
  result := true;

  Fic_Bidon := Rep_Ouvrage +'\SousDetail.xls';

  TRY
    if not FileExists (Fic_Bidon) then
    begin
      if not CreateTheFichierXls (Fic_Bidon) then exit;
    end;
    if fileExists ('c:\PGI00\STD\MetrePgi.xla') Then
       Begin
	     // Copie de la macro dans le répertoire d'office
       AddVbModuleFromFile ('c:\PGI00\STD\','MetrePgi.xla');
       //Execution d'excel avec le fichier bidon.xls
       FileExecAndWait (GetExcelPath + 'EXCEL.exe "' + Fic_Bidon + '"');
       end
    else
       Begin
       PGIBoxAF (traduirememoire('Le Fichier MetrePGI.xla n''existe pas'),traduireMemoire('METRE'));
       end;
  FINALLY
    // suppression du fichier des variables
    if FileExists ('C:\pgi00\STD\Variables.txt') then
       DeleteFichier('C:\pgi00\STD\Variables.txt');
    // Suppression de la Macro dans le répertoire d'office
    DelVBModule ('MetrePgi.xla');
    //listage des fichiers du répertoire Ouvrage et
    //suppression de ceux-ci + le répertoire
    if FileExists (Fic_Bidon) then
       DeleteFichier (Fic_Bidon);
    EpureRepertoireOuvrage;
    //
  END;

end;

// Gestion des metres bibliothèque dans le cadre d'un artcile
procedure TStockMetre.DocMetreBiblio(TOBL:TOB);
Var FicArt: String;
    CodeArt: String;
    TypeArt: String;
    //MetreArticle: TMetreArt;
    Prefixe: string;
    Q_Variables : TQuery;
begin

  if TOBL.NomTable = 'LIGNE' then
    Prefixe := 'GL_'
  Else if TOBL.NomTable = 'LIGNEOUV' Then
    Prefixe := 'BLO_'
  Else
    Exit;

  TypeArt := TOBL.GetValue(Prefixe + 'TYPEARTICLE');
  CodeArt := TOBL.GetValue(Prefixe + 'ARTICLE');

  // Si c'est un ouvrage ou une nomenclature par de gestion ici
  if (TypeArt = 'NOM') or (TypeArt = 'OUV') then Exit;

  // Recherche du Nom du fichier Article de type XLS
  MetreArticle := TMetreArt.create;
  FicArt := MetreArticle.GesDoc(CodeArt,TypeArt);
  MetreArticle.free;

  if ficart = '' then
     Begin
     // Lecture de la table des Variables
     Q_Variables := OpenSQL('SELECT * FROM BVARIABLES WHERE BVA_TYPE = "B" '
                          + 'AND BVA_ARTICLE = "' + CodeArt + '"', true);
     if Q_Variables.EOF then
        Begin
        Ferme(Q_Variables);
        exit;
        end
     else
        Begin
        Ferme(Q_Variables);
        RecupMetreBiblio(TOBL, FicArt);
        EcrireText();
        exit;
        end;
     end
  else
     Begin
     RecupMetreBiblio(TOBL, FicArt);
     EcrireText();
     end;

end;

// Epuration du repertoire ouvrage après traitement des différents fichiers
procedure TStockMetre.EpureRepertoireOuvrage;
var NomFic   : string;
    Nomfic_2 : string;
    Rec : TSearchRec;
begin

  Nomfic:= Rep_Ouvrage + '\*.*';
  Nomfic_2 := fRepertBase + '\' + fRepert + '\';

  if FindFirst (Nomfic,faAnyFile,Rec) = 0 then
     begin
     if (rec.name <> '.') and (rec.name <> '..') then
        begin
            CopieFichier(rep_ouvrage + '\'+ Rec.Name, NomFic_2 + Rec.Name);
            deleteFichier (rep_ouvrage + '\'+ rec.Name);
        end;
     while FindNext (Rec) = 0 do
        begin
        if (rec.name <> '.') and (rec.name <> '..') then
           begin
              CopieFichier(rep_ouvrage + '\'+ Rec.Name, NomFic_2 + Rec.Name);
              deleteFichier (Rep_Ouvrage +'\'+ Rec.Name);
           end;
        end;
     end;

  FindClose (Rec);

  // Suppression repertoire Ouvrage
  deleteRepertoire(Rep_Ouvrage);

end;

Procedure TStockMetre.LanceVar(Action : TactionFiche);
begin

  //Appel de la Fenêtre de saisie des Variables du Document
  //ChargeVariables(Action, TobSaisieMetre, SNaturePiece, SSouche, SNumeroPiece, SInd);

end;
{***********A.G.L.***********************************************
Auteur  ...... : franck Vautrain
Créé le ...... : 05/09/2003
Modifié le ... : 08/09/2003
Description .. : Ecriture des Variables dans un fichier Txt
Mots clefs ... : ECRIRE VARIABLES.TXT
*****************************************************************}
procedure TStockMetre.EcrireText();
var
  //on crée une variable Textfile
  fichier   : textfile;
  Rep_Metre : string;
  Libelle   : string;
  Valeur    : String;
  W         : Integer;
begin

  if TobSaisieMetre = nil then exit;

  //Appel de la Fenêtre de saisie
  if OkSaisieVariables() then
  Begin
  	//if not Saisievariables(TobSaisieMetre) then exit;
  end;

  //Recupération base de travail métrés
  Rep_Metre := 'C:\pgi00\STD\';

  if not DirectoryExists(Rep_Metre) then
  Creationdir(Rep_Metre);

  if TobSaisieMetre.detail.count > 0 then
  begin
    For W:=0 To TobSaisieMetre.detail.count-1 do
    Begin
      Libelle:=TobSaisieMetre.detail[W].getValue('BVD_CODEVARIABLE');
      Valeur:=TobSaisieMetre.detail[W].getValue('BVD_VALEUR');
      if not ControleText(Libelle) then
      Begin
        //Chargement du nom du fichier
                AssignFile(fichier , Rep_Metre + 'Variables.txt');
        //Accès en Lecture/écriture sur le fichier
        FileMode := 2;
                if FileExists (Rep_Metre + 'Variables.txt') then
        	Append(fichier)
        else
        	Rewrite(fichier);
        if valeur = '' then
        Begin
          //ajoute une ligne au fichier et l'écrit Physiquement
          Writeln(fichier,Libelle);
          Flush(fichier);
        end
        else
        Begin
          //ajoute une ligne au fichier et l'écrit Physiquement
          Writeln(fichier,Libelle + ';' + Valeur);
          Flush(fichier);
        end;
        //ferme le fichier
        CloseFile(fichier);
      end;
    end;
  end;

end;

Function TStockMetre.OkSaisieVariables() : Boolean;
var Val : String;
    W   : Integer;
Begin

  result := false;

  if TobSaisieMetre = nil then exit;

  if TobSaisieMetre.detail.count > 0 then
     begin
      For W:=0 To TobSaisieMetre.detail.count-1 do
        Begin
          Val := TobSaisieMetre.detail[W].getValue('BVD_VALEUR');
          if Val = '' then
            begin
              result := true;
              break;
            end;
        end;
     end;

End;

// Controle de la duplicité des variables dans le cadre d'un Ouvrage principalement
function TStockMetre.ControleText(Libelle : string): Boolean;
Var F         : TextFile;
    Ligne     : String;
//    RepertTmp : String;

Begin

  Result := False;

  if Not FileExists ('C:\pgi00\STD\Variables.txt') then
     exit;

  AssignFile(F, 'C:\pgi00\STD\Variables.txt');

  reset(F);

  while not SeekEof(F) do
    begin
      Readln(F, Ligne);
      Ligne := ExtraireTexte(ligne, Libelle, ';', 0, 0, length(Ligne));
      Ligne := replace(ligne, ';', '');
      if Ligne = Libelle then
        Begin
          Result := True;
          CloseFile(F);
          Exit;
        end;
    end;

  CloseFile(F);

end;


//permet d'extraire une chaine de caractère d'une chaine de caractère en
//fonction d'un caractère spécial
function TStockMetre.ExtraireTexte(Source, TexteDep, Limites:string; Depart, Deb, Longueur:integer):string;
var
a,b:integer;
Fin:integer;
TmpS:string;
BVal:boolean;
begin
  Result:='';

  //faut ki est kelke chose sinon ca va pas ;)
  if Source='' then exit;

  //cherche TexteDep dans Source
  a:=Depart;
  BVal:=true;

  while BVal do begin
    //dès k'il trouve
    if copy(Source,a,Length(TexteDep))=TexteDep then
      begin
        BVal:=false;//arrête le boucle
        Depart:=a;//change le dépat
      end;

    //si c a la fin, ben alors il a rien trouvé, terminé
    if (a>=Length(Source)) and (BVal=true) then
      begin
        exit;
      end;
    inc(a,1);//incrémente
end;


//Cherche le début (Deb)
a:=Depart;
BVal := true;
while BVal=true do begin
    //cherche un caractère limite
    TmpS := copy(Source,a,1);
    for b:=1 to Length(Limites) do begin
    if copy(Limites,b,1)=TmpS then begin
        BVal:=false;//y'en a un, on a trouvé le début
        a:=a+2;
    end;
    end;
    dec(a,1);
    if a<1 then begin
    BVal:=false;//on est au bout :)
    a:=a+1;
    end;
end;
Deb:=a;

//Cherche la fin (Fin)
a:=Depart;
BVal:=true;
while BVal do begin
    //cherche un caractère limite
    TmpS := copy(Source,a,1);
    for b:=1 to Length(Limites) do begin
    if copy(Limites,b,1)=TmpS then begin
        BVal:=false;//y'en a un, on a trouvé le début
        a:=a-1;
    end;
    end;
Inc(a,1);
    if TmpS='' then begin
    BVal:=false;//on est au bout :)
    a:=a-1;
    end;
end;
Fin:=a;

Longueur := Fin-Deb;
//prend le morceu
Result := copy(Source,Deb,Longueur);

end;

{***********A.G.L.***********************************************
Auteur  ...... : franck Vautrain
Créé le ...... : 05/09/2003
Modifié le ... : 08/09/2003
Description .. : Ecriture des Variables dans la Table BVARDOC
Mots clefs ... : ECRIRE BVARDOC
*****************************************************************}
procedure TStockMetre.CreationVarDoc(CodeArt, NaturePiece, Souche: string;
                                     IO, NumeroPiece, I : Integer);

Var
  Q_Variables : TQuery;
  IndexFile   : Integer;
  W           : Integer;
  TOBFille    : TOB;
begin

  IndexFile := IO;

  // Lecture de la table des Variables
  Q_Variables := OpenSQL('SELECT * FROM BVARIABLES WHERE BVA_TYPE = "B" '
                       + 'AND BVA_ARTICLE = "' + CodeArt + '"', true);

  if Q_variables.EOF THEN BEGIN Ferme (Q_Variables) ; Exit; END;

  TOBVariables := TOB.Create('Var Article', nil, -1);

  // Creation de toutes les TOB dans la propriétés détails en fonction du Quéry
  TOBVariables.LoadDetailDB('BVARIABLES', '', '', Q_variables, false);
  Ferme (Q_Variables);

  if TobVariables <> nil then
    if TobVariables.detail.count > 0 then
       begin
        For W:=0 To TobVariables.detail.count-1 do
          Begin
            Tobfille := TOBSaisieMetre.Findfirst(['BVD_CODEVARIABLE'],
                                                 [TOBVariables.detail[W].getValue('BVA_CODEVARIABLE')],
                                                 true);
            if TOBFille = nil then
            begin
               // Chargement d'une grille avec les enregistrement de la Tob
               TobFille := Tob.Create('BVARDOC', TOBSaisieMetre, -1);
               AddChampsSupVar (TOBFille);
               TobFille.Putvalue('BVD_NATUREPIECE', NaturePiece);
               TobFille.Putvalue('BVD_SOUCHE',Souche);
               TobFille.Putvalue('BVD_NUMERO',IntToStr(NumeroPiece));
               TobFille.Putvalue('BVD_INDICE',IntToStr(I));
               TobFille.Putvalue('BVD_INDEXFILE',IntToStr(Indexfile));
               TobFille.Putvalue('BVD_CODEVARIABLE',TobVariables.detail[W].getValue('BVA_CODEVARIABLE'));
               TobFille.Putvalue('BVD_LIBELLE',TobVariables.detail[W].getValue('BVA_LIBELLE'));
               TobFille.Putvalue('BVD_VALEUR',TobVariables.detail[W].getValue('BVA_VALEUR'));
            end;
          End;
       end;

  TobVariables.free;

end;

function TStockMetre.RecupVariableInFic(Nom_Fic_Text: string): variant;
var F : Textfile;
    Ligne : string;
begin

  Result := False;

  AssignFile(F, Nom_Fic_Text);

  {$I-}
  reset(F);
  {$I+}

  while not SeekEof(F) do
    begin
      Readln(F, Ligne);
      Result := Ligne;
    end;

  CloseFile(F);

  if IsNumeric(result) then
     result := Arrondi(VarasType(Result,vardouble),V_PGI.OkDecQ)
  else
     Begin
       Result := 0;
       //Mis en commentaire par BRL 19/03 suite pb si pas de métré en sous-détail
       //Il faudrait géré le cas et afficher un message spécifiant le sous-détail concerné
       //PGIBox(TraduireMemoire('La valeur n''est pas numérique'),TraduireMemoire('METRES'));
     End;

end;

//-- Chargement TOB Variables pour gestion des metrés
procedure TStockMetre.ChargementVariables(CleDoc : R_CleDoc);
Var
  Q  : TQuery;
  I  : Integer;
  NumDoc : String;
  TOBFille : TOB;
begin

  NumDoc := CleDoc.NaturePiece + '_' + CleDoc.Souche + '_' +
            IntToStr(CleDoc.NumeroPiece) + '_' + IntToStr(CleDoc.Indice);

  SNaturePiece := CleDoc.NaturePiece;
  SSouche := CleDoc.Souche;
  SNumeroPiece := CleDoc.NumeroPiece;
  SInd := CleDoc.Indice;

  //Creation des TOB Virtuelles Variable Application, Variables Document et
  //Variables déjà renseignées
  TOBVariables := TOB.Create('Var Application', nil, -1);
  TRY
  //Lecture de la table des variables Application
  Q := OpenSQL('SELECT * FROM BVARIABLES WHERE BVA_TYPE = "A"', true);
  // Creation de toutes les TOB dans la propriétés détails en fonction du Quéry
  TOBVariables.LoadDetailDB('BVARIABLES', '', '', Q, False);
  Ferme (Q);

  //Lecture de la table des variables documents
  Q := OpenSQL('SELECT * FROM BVARDOC WHERE BVD_NATUREPIECE = "' + CleDoc.NaturePiece
                                        + '" AND BVD_SOUCHE="' + CleDoc.Souche
                                        + '" AND BVD_NUMERO= ' + IntToStr(CleDoc.NumeroPiece)
                                        + '  AND BVD_INDICE= ' + IntToStr(CleDoc.Indice), true);

  // Creation de toutes les TOB dans la propriétés détails en fonction du Quéry
  TOBSaisieMetre.LoadDetailDB('BVARDOC', '', '', Q, true);
  Ferme (Q);
  for I := 0 to TOBSaisiemetre.detail.count -1 do
  begin
  	AddChampsSupVar (TOBSaisiemetre.detail[I]);
  end;

  if TOBSaisieMetre.detail.count > 0 then exit;

  //Chargement des enregistrements de TOBVariables vers TOBSaisieMetre
  if TobVariables <> nil then
    if TobVariables.detail.count > 0 then
       begin
        For I:=0 To TobVariables.detail.count-1 do
          Begin
            // Chargement d'une grille avec les enregistrement de la Tob
            TobFille := Tob.Create('BVARDOC', TOBSaisieMetre, -1);
            AddChampsSupVar (TOBFille);
            TobFille.Putvalue('BVD_NATUREPIECE', CleDoc.NaturePiece);
            TobFille.Putvalue('BVD_SOUCHE',CleDoc.Souche);
            TobFille.Putvalue('BVD_NUMERO',IntToStr(CleDoc.NumeroPiece));
            TobFille.Putvalue('BVD_INDICE',IntToStr(CleDoc.Indice));
            TobFille.Putvalue('BVD_INDEXFILE',IntToStr(0));
            TobFille.Putvalue('BVD_CODEVARIABLE',TobVariables.detail[I].getValue('BVA_CODEVARIABLE'));
            TobFille.Putvalue('BVD_LIBELLE',TobVariables.detail[I].getValue('BVA_LIBELLE'));
            TobFille.Putvalue('BVD_VALEUR',TobVariables.detail[I].getValue('BVA_VALEUR'));
          End;
       End;
  FINALLY
    TobVariables.free;
  END;
end;

//Vérification si les metrés sont utilisables en saisie
function TStockMetre.ControlMetre: boolean;
var Dest : string;
Begin

    result := true;

    // controle si macro excel existe
    if not fileExists('c:\PGI00\STD\MetrePgi.xla') then
      begin
        Result := false;
        Exit;
      end;

    // Controle si excel installé sur la machine
    dest := GetOfficeDir + 'XlStart\';
    if not DirectoryExists(dest) then
      begin
        Result := false;
        exit;
      end;

    //Controle du parametre de gestion des métrés
    if not GetParamSoc('SO_BTMETREDOC') then
      begin
        Result := false;
        exit;
      end;

End;

//Contrôle si la nature de piéce peut gérer ou non les métrés
Function TStockMetre.AutorisationMetre(NatPiece : string) : Boolean;
Begin

	result := false;

  if NatPiece = GetParamSoc('SO_AFNATAFFAIRE') then
	  result := True
  Else if NatPiece = GetParamSoc('SO_AFNATPROPOSITION') then
	  result := True
  Else if (Pos(NatPiece,'FBT;FBP')>0) then
	  result := True
  Else if NatPiece = 'AFF' then
 	  result := True
  Else if NatPiece = 'DAP' then
 	  result := True;

end;

procedure TStockMetre.AddChampsSupVar(TOBFille: TOB);
begin
	TOBFille.AddChampSupValeur('SEL','');
end;

{$IFDEF CEAGLCLIENT}
procedure TStockMetre.SetTOBMetres (TOBmetre : TOB);
var Indice : integer;
begin
	sTOBMetres.clearDetail;
	if TOBmetre.detail.count = 0 then exit;
	repeat
  	TOBmetre.detail[Indice].ChangeParent(sTOBMetres,-1);
  until Indice > TOBMetre.detail.count -1;
end;
{$ENDIF}

end.
