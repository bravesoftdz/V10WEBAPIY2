unit EtudesUtil;

interface

uses  UTOB,
      Hctrls,
      ParamSoc,
      Dicobtp,
      FileCtrl,
      SYSUtils,
      ENtGC,
      FactComm,
      HEnt1
{$IFDEF BTP}
     ,BTPUtil
{$ENDIF}
		,uEntCommun
     ;
Type TActionTrait = (TatNormal,TatAccept,TatBordereaux,TatTout);
procedure AfficheLaLigne (GS : THGrid;TOBEtude : TOB; lescolonnes: string ;Arow : integer);
procedure AfficheLagrille (GS : THGrid;TOBEtude:TOB;Lescolonnes: string;Arow : integer);
procedure AddEtudeAttachP (TOBEtude,TOBPrinc,TOBLastP,TOBNat : TOB; ARow : integer;NatureAuxi : string='CLI');
procedure AddEtudeInitiale (TOBEtude,TOBNATD: TOB;Traitement : TActionTrait;NatureAuxi : string='CLI');
procedure AddEtudePrincipale (TOBEtude,TOBLastP,TobNat: TOB;Arow : integer; Traitement : TactionTrait; NatureAuxi : string='CLI');
procedure AddEtudeSecond (TOBEtude,TOBPrinc,TOBLastP,TOBNat : TOB; ARow : integer;NatureAuxi : string='CLI');
function ajouteTobstructure (Pere: TOB;Indice: integer):TOB;
procedure AddChampsSupEtude (TOBEtude : TOB);
procedure ChampSupStructure (TOBR : TOB;touteslesfilles: boolean);
procedure FaireMenagePiece (TOBorig,TOBcmp,TOBAnnul : TOB);
function FindLastPrincipal (TOBEtude:tob; Traitement : TActionTrait=TatNormal ):TOB;
function FindLastSecond (TOBEtude,TOBPrinc: TOB;CodeNature:string;var Arow : integer):TOB;
function FindNatureEtude (TOBNatD : TOB; chaine : string): TOB;
function findPrincipal (TOBetude : TOB;Arow : integer) : TOB;
function GetTobEtude (TOBEtudes : TOB; Arow : integer) : TOB;
function InitTobEtude (TOBEtude,TOBNat,TOBL : TOB;Libelle,NatureAuxi : string): boolean;
procedure InitBordereauClient (TOBEtude : Tob);
procedure NettoyageAppelOffre (TOBEtudes : TOB);

function recupNomfic (nomfic : string): string;
procedure FaireMenageRepertoire (TOBEtude : TOB; Traitement : TActionTrait = TatTout);
//function CopieFichier (nomOrig,Nomdest : string):boolean;
//function EnregistreFichier (TypeDoc : string;TobEtude:Tob;Fichini,FicOut : string) : string;
function ExisteFichier (Fichier : string; TOBEtude : TOB):boolean;

implementation

function FindNaturePrincipale (TOBNatD : TOB; Action : TactionTrait ): TOB;
begin
result := nil;
if TOBNatd.detail.count = 0 then exit;
if Action = TatBordereaux then result := TOBNatd.FindFirst (['CO_LIBRE'],['PRINC1'],true)
													ELSE result := TOBNatd.FindFirst (['CO_LIBRE'],['PRINC'],true);
end;

function strdelete (chaine,sousChaine: string): string;
var indice : integer;
begin
repeat
  indice := pos (souschaine,chaine);
  if Indice = 1 then chaine := copy (chaine,2,255)
  else if Indice > 1 then Chaine := copy (chaine,1,Indice -1)+copy(chaine,Indice+1,255);
until indice <= 0;
result := Chaine;
end;

procedure AddEtudeInitiale (TOBEtude,TOBNATD: TOB;Traitement : TActionTrait;NatureAuxi : string='CLI');
var TOBL,TOBNat : TOB;
begin
TOBNAt := FindNaturePrincipale (TOBNATD,Traitement);
if TOBNAT = nil then exit;
TOBL := TOB.Create ('BDETETUDE',TOBEtude,-1);
InitTobEtude (TOBEtude,TOBNat,TOBL,TOBNat.getValue('CO_LIBELLE'),NatureAuxi);
TOBL.addchampsupValeur('NEW','X');
TOBL.PutValue('BDE_ORDRE',1);
if Traitement = TatBordereaux Then TOBL.PutValue('BDE_QUALIFNAT','PRINC1')
															else TOBL.PutValue('BDE_QUALIFNAT','PRINC');
TOBL.PutValue('BDE_SELECTIONNE','X');
end;

procedure AddEtudeAttachP (TOBEtude,TOBPrinc,TOBLastP,TOBNat : TOB; ARow : integer;NatureAuxi : string='CLI');
var TOBL : TOB;
begin
TOBL := TOB.Create ('BDETETUDE',TOBEtude,Arow-1);
InitTobEtude (TOBEtude,TOBNat,TOBL,TOBNat.getValue('CO_LIBELLE'),NatureAuxi);
TOBL.addchampsupValeur('NEW','X');
TOBL.PutValue('BDE_ORDRE',TOBPrinc.getValue('BDE_ORDRE'));
TOBL.PutValue('BDE_NATUREDOC',TOBNat.getValue('CO_CODE'));
if TOBLastp = nil then TOBL.PutValue ('BDE_INDICE',1)
                  else TOBL.putValue('BDE_INDICE',TOBLastp.getValue('BDE_INDICE')+1);
TOBL.putVALUE('BDE_PIECEASSOCIEE',TOBPrinc.GetValue('BDE_PIECEASSOCIEE')); // rajout apres creation doc principal
TOBL.PutValue('BDE_NATUREPIECEG',TOBPrinc.GetValue('BDE_NATUREPIECEG'));
TOBL.PutValue('BDE_SOUCHE',TOBPrinc.GetValue('BDE_SOUCHE'));
TOBL.PutValue('BDE_NUMERO',TOBPrinc.GetValue('BDE_NUMERO'));
TOBL.PutValue('BDE_INDICEG',TOBPrinc.GetValue('BDE_INDICEG'));
TOBL.PutValue('BDE_QUALIFNAT','PRINC1');
end;

procedure AddEtudePrincipale (TOBEtude,TOBLastP,TobNat: TOB;Arow : integer; Traitement : TactionTrait; NatureAuxi : string='CLI');
var TOBL : TOB;
begin
TOBL := TOB.Create ('BDETETUDE',TOBEtude,Arow);
InitTobEtude (TOBEtude,TOBNat,TOBL,TOBNat.getValue('CO_LIBELLE'),NatureAuxi);
TOBL.addchampsupValeur('NEW','X');
if TOBLastP = nil then TOBL.PutValue('BDE_ORDRE',1)
                  else TOBL.PutValue('BDE_ORDRE',TOBLastp.getValue('BDE_ORDRE')+1);
if Traitement <> TatBordereaux then TOBL.PutValue('BDE_QUALIFNAT','PRINC')
															 else TOBL.PutValue('BDE_QUALIFNAT','PRINC1');
TOBL.PutValue('BDE_SELECTIONNE','X');
end;

procedure AddEtudeSecond (TOBEtude,TOBPrinc,TOBLastP,TOBNat : TOB; ARow : integer;NatureAuxi : string='CLI');
var TOBL : TOB;
begin
TOBL := TOB.Create ('BDETETUDE',TOBEtude,Arow-1);
InitTobEtude (TOBEtude,TOBNat,TOBL,TOBNat.getValue('CO_LIBELLE'),NatureAuxi);
TOBL.addchampsupValeur('NEW','X');
TOBL.PutValue('BDE_ORDRE',TOBPrinc.getValue('BDE_ORDRE'));
TOBL.PutValue('BDE_NATUREDOC',TOBNat.getValue('CO_CODE'));
if TOBLastp = nil then TOBL.PutValue ('BDE_INDICE',1)
                 else TOBL.putValue('BDE_INDICE',TOBLastp.getValue('BDE_INDICE')+1);
TOBL.PutValue('BDE_QUALIFNAT','SECOND');
end;

procedure AfficheLaLigne (GS : THGrid;TOBEtude : TOB; lescolonnes: string ;Arow : integer);
var TOBL : TOB;
begin
TOBL := GetTOBEtude (TOBEtude,Arow);
if TOBL = nil then exit;
TOBL.PutLigneGrid (GS,Arow,false,false,Lescolonnes);
GS.InvalidateRow (Arow);
end;

procedure AfficheLagrille (GS : THGrid;TOBEtude:TOB;Lescolonnes: string;Arow : integer);
var indice : integer;
begin
for Indice := Arow to GS.rowcount -1 do
    begin
    AfficheLaLigne (GS,TOBEtude,LesColonnes,Indice);
    end;
GS.Refresh;
end;

function ajouteTobstructure (Pere: TOB;Indice: integer):TOB;
var TOBR: TOB;
begin
TOBR := TOB.create ('BSTRDOC',pere,-1);
ChampSupStructure (TOBR,false);
result := TOBR;
end;

procedure ChampSupStructure (TOBR : TOB;touteslesfilles: boolean);
begin
end;

procedure AddChampsSupEtude (TOBEtude : TOB);
begin
	TOBEtude.addchampsupValeur ('LAFF_NOMCLI','');
	TOBEtude.addchampsupValeur ('NATUREAUXI','');
end;

(*
function CopieFichier (nomOrig,Nomdest : string):boolean;
var
FromF, ToF: file;
NumRead, NumWritten: Integer;
Buf: array[1..2048] of Char;
begin
AssignFile(FromF, NomOrig);
{$I-}
Reset(FromF, 1);
{$I+}
if IoResult = 0 then
  begin
  AssignFile(ToF, Nomdest);
  {$I-}
  Rewrite(ToF, 1);
  {$I+}
  if IoResult = 0 then
     begin
     result := true;
     repeat
         BlockRead(FromF, Buf, SizeOf(Buf), NumRead);
         BlockWrite(ToF, Buf, NumRead, NumWritten);
     until (NumRead = 0) or (NumWritten <> NumRead);
     end else Result := false;
  CloseFile(ToF);
  end else result := false;
CloseFile(FromF);
end;

function EnregistreFichier (TypeDoc : string;TobEtude:Tob;Fichini,FicOut : string) : string;
var repertoire,nomfic,nomOut : string;
begin
result := '';

repertoire := GetParamSoc('SO_BTREPAPPOFF');
if repertoire = '' then
   begin
   PGIBoxAF ('Veuillez définir le répertoire de stockage dans les paramètres sociétés','PARAM');
   exit;
   end;
   nomfic := recupNomfic (Fichini);
   nomOut := recupNomfic (FicOut);
   repertoire := repertoire +'\'+TobEtude.GetValue('AFF_AFFAIRE');
   if not DirectoryExists (repertoire) then
   begin
      if not Createdir (repertoire) then exit;
   end;
   if CopieFichier (Fichini,repertoire+'\'+nomOut) then result := Repertoire+'\'+nomOut;
end;
*)

function ExisteFichier (Fichier : string; TOBEtude : TOB):boolean;
var Indice : integer;
    TOBL : TOB;
    ficCmp,FicRech : string;
begin
result := false;
if TOBEtude = nil then exit;
if TOBEtude.detail.count = 0 then exit;
FicRech := RecupNomfic (fichier);
for Indice := 0 to TOBEtude.detail.count -1 do
    begin
    TOBL := TOBEtude.detail[Indice];
    if TOBL = nil then break;
    if TOBL.GetValue('BDE_NAME') = '' then continue;
    ficCmp := RecupNomfic(TOBL.GetValue('BDE_NAME'));
    if FicCmp = FicRech then BEGIN result := true; Break; END;
    end;
end;

procedure FaireMenagePiece (TOBorig,TOBcmp,TOBAnnul : TOB);
var Indice : integer;
    TOBL,TOBF : TOB;
    CleDoc : R_Cledoc;
begin

  if TOBOrig <> nil then
  BEGIN
    for Indice:=0 to TOBOrig.detail.count -1 do
    begin
      TOBL := TOBOrig.detail[Indice];
      if TOBL.getValue('BDE_PIECEASSOCIEE') <> '' then
      begin
        DecodeRefPiece (TOBL.GetValue('BDE_PIECEASSOCIEE'),CleDoc);
        TOBF := TOBCmp.findfirst(['BDE_PIECEASSOCIEE'],[TOBL.GetValue('BDE_PIECEASSOCIEE')],true);
        {$IFDEF BTP}
        if (TOBF = nil) (* and (TOBL.getValue('BDE_QUALIFNAT')<>'PRINC1') *) then
        begin
          BTPSupprimePiece (cledoc);
          BTPSupprimePieceFrais (cledoc);
        end;
        {$ENDIF}
      end;
    end;
	END;

  if TOBannul <> nil then
  BEGIN
    for Indice:=0 to TOBAnnul.detail.count -1 do
    begin
      TOBL := TOBAnnul.detail[Indice];
      if TOBL.getValue('BDE_PIECEASSOCIEE') <> '' then
      begin
        DecodeRefPiece (TOBL.GetValue('BDE_PIECEASSOCIEE'),CleDoc);
        //          if (TOBL.getValue('BDE_QUALIFNAT')<>'PRINC1') then
        begin
          BTPSupprimePiece (cledoc);
          BTPSupprimePieceFrais (cledoc);
        end;
      end;
    end;
  END;
end;

procedure FaireMenageRepertoire (TOBEtude : TOB; Traitement : TActionTrait = TatTout);

          procedure deleteRepert (repertoire : string);
          var nomfic : string;
          Rec : TSearchRec;
          begin
          Nomfic:= Repertoire+'\*.*';
          if FindFirst (Nomfic,faAnyFile,Rec) = 0 then
             begin
             if (rec.name <> '.') and (rec.name <> '..') then deleteFile (repertoire+'\'+rec.Name);
             while FindNext (Rec) = 0 do
                 begin
                 if (rec.name <> '.') and (rec.name <> '..') then deleteFile (repertoire+'\'+rec.Name);
                 end;
             end;
          FindClose (Rec);
          removedir(Repertoire);
          end;

var repertoire,Nomfic : string;
    Rec : TSearchRec;
begin

  if Traitement = TatTout then
  begin
  	FaireMenageRepertoire (TOBEtude,TatNormal);
  	FaireMenageRepertoire (TOBEtude,TatBordereaux);
    exit;
  end;

  if Traitement <> TatBordereaux then repertoire := GetParamSoc('SO_BTREPAPPOFF')
                     						 else repertoire := GetParamSoc('SO_BTREPBORDPRIX');

if Repertoire = '' then exit;
if TOBEtude.GetValue('AFF_AFFAIRE')<> '' then
begin
	repertoire := repertoire + '\' + TOBEtude.GetValue('AFF_AFFAIRE');
end else
begin
	repertoire := repertoire + '\T' + TOBEtude.GetValue('AFF_TIERS');
end;
Nomfic:= Repertoire+'\*.*';
if FindFirst (Nomfic,faAnyFile,Rec) = 0 then
   begin
   if (rec.name <> '.') and (rec.name <> '..') then
      begin
      if not ExisteFichier (Rec.name,TOBEtude) then deleteFile (repertoire+'\'+rec.Name);
      end;
   while FindNext (Rec) = 0 do
       begin
       if (rec.name <> '.') and (rec.name <> '..') then
          begin
          if not ExisteFichier (Rec.name,TOBEtude) then deleteFile (repertoire+'\'+rec.Name);
          end;
       end;
   end;
FindClose (Rec);
if TOBEtude.detail.count = 0 then
   begin
   deleteRepert (repertoire);
   end;
end;

function FindLastPrincipal (TOBEtude:tob; Traitement : TActionTrait=TatNormal ):TOB;
var TOBL : TOB;
    Indice : integer;
    TheRecherche : string;
begin
if Traitement = TatBordereaux then TheRecherche := 'PRINC1'
															else TheRecherche := 'PRINC';
Result := nil;
for Indice:= 0 to TOBEtude.detail.count -1 do
    begin
    TOBL := TOBEtude.detail[Indice];
    if (TOBL.GetValue('BDE_QUALIFNAT')= TheRecherche) then
       begin
       Result := TOBL;
       end;
    end;
end;

function FindLastSecond (TOBEtude,TOBPrinc: TOB;CodeNature:string;var Arow : integer):TOB;
var TOBL : TOB;
    Indice : integer;
begin
Result := nil;
Arow:= -1;
for Indice:= 0 to TOBEtude.detail.count -1 do
    begin
    TOBL := TOBEtude.detail[Indice];
    if (TOBL.GetValue('BDE_ORDRE')> TOBPrinc.getValue('BDE_ORDRE')) then BEGIN Arow := Indice; break; END;
    if (TOBL.GetValue('BDE_QUALIFNAT')<>'PRINC') and (TOBL.GetValue('BDE_ORDRE') = TOBPrinc.getValue('BDE_ORDRE')) then
       if (TOBL.getValue('BDE_NATUREDOC')> CodeNature) then BEGIN Arow := Indice; break; END;
    if (TOBL.GetValue('BDE_ORDRE')= TOBPrinc.getValue('BDE_ORDRE')) and
       (TOBL.GetValue('BDE_NATUREDOC')= CodeNature) then
       begin
       Result := TOBL;
       Arow := Indice+1;
       end;
    end;
end;

function FindNatureEtude (TOBNatD : TOB; chaine : string): TOB;
begin
result := nil;
if TOBNatd.detail.count = 0 then exit;
chaine := StrDelete (Chaine,'&');
result := TOBNatd.FindFirst (['CO_LIBELLE'],[chaine],true);
end;

function findPrincipal (TOBetude : TOB;Arow : integer) : TOB;
var TOBLoc : TOB;
begin
result := nil;
if TOBEtude.detail.count = 0 then exit;
TObLoc := TOBEtude.findFirst (['BDE_ORDRE','BDE_QUALIFNAT'],[TobEtude.detail[Arow-1].GetValue('BDE_ORDRE'),'PRINC'],true);
result := TobLoc;
end;

function GetTobEtude (TOBEtudes : TOB; Arow : integer) : TOB;
begin
Result:=Nil ;
if ((ARow<=0) or (ARow>TobEtudes.Detail.Count)) then Exit ;
result := TOBEtudes.Detail[Arow-1];
end;


function InitTobEtude (TOBEtude,TOBNat,TOBL : TOB;Libelle,NatureAuxi : string): boolean;
begin
result := true;
if TOBNat = nil then
   begin
   result := false;
   exit;
   end;
TOBL.PutValue ('BDE_CLIENT',TOBEtude.getValue('AFF_TIERS'));
TOBL.PutValue ('BDE_NATUREAUXI',NatureAuxi);
TOBL.PutValue ('BDE_AFFAIRE',TOBEtude.getValue('AFF_AFFAIRE'));
TOBL.PutValue ('BDE_NATUREDOC',TOBNat.GetValue('CO_CODE'));
TOBL.PutValue ('BDE_DESIGNATION',strdelete (TOBNat.GetValue('CO_LIBELLE'),'&'));
TOBL.PutValue ('BDE_DATEDEPART',StrToDate(DateToStr(Date) ));
TOBL.PutValue ('BDE_DATEFIN',Idate2099 );
end;

procedure InitBordereauClient (TOBEtude : Tob);
BEGIN
	if not TOBEtude.FieldExists ('AFF_TIERS') then
  begin
    TOBEtude.AddChampSupValeur('AFF_TIERS','');
    TOBEtude.AddChampSupValeur('AFF_AFFAIRE','');
    TOBEtude.AddChampSupValeur('AFF_AFFAIRE0','');
    TOBEtude.AddChampSupValeur('AFF_AFFAIRE1','');
    TOBEtude.AddChampSupValeur('AFF_AFFAIRE2','');
    TOBEtude.AddChampSupValeur('AFF_AFFAIRE3','');
    TOBEtude.AddChampSupValeur('AFF_AVENANT','');
    TOBEtude.AddChampSupValeur('AFF_LIBELLE','');
    TOBEtude.AddChampSupValeur('LAFF_NOMCLI','');
    TOBEtude.AddChampSupValeur('NATUREAUXI','');
  end else
  begin
    TOBEtude.PutValue('AFF_TIERS','');
    TOBEtude.PutValue('AFF_AFFAIRE','');
    TOBEtude.PutValue('AFF_AFFAIRE0','');
    TOBEtude.PutValue('AFF_AFFAIRE1','');
    TOBEtude.PutValue('AFF_AFFAIRE2','');
    TOBEtude.PutValue('AFF_AFFAIRE3','');
    TOBEtude.PutValue('AFF_AVENANT','');
    TOBEtude.PutValue('AFF_LIBELLE','');
    TOBEtude.PutValue('LAFF_NOMCLI','');
    TOBEtude.PutValue('NATUREAUXI','');
  end;
END;

function recupNomfic (nomfic : string): string;
var position : integer;
begin
position := length (nomfic);
repeat
  if copy(nomfic,position,1) = '\' then break;
  dec (position);
until position <= 1;
if position > 1 then result := copy (nomfic,position+1,length(nomfic)-position+1)
                else Result := nomfic;
end;

procedure NettoyageAppelOffre (TOBEtudes : TOB);
var
  TOBLOC: TOB;
  Indice: Integer;
begin
  if TOBEtudes = nil then exit;
  TOBLOC := TOB.Create ('La Locale', nil, -1) ; //mcd 28/01/03 chgmt nom tob
  TOBLOC.dupliquer (TOBEtudes, false, true) ;
  FaireMenageRepertoire (TOBLOC) ;
  TOBLOC.free;
  FaireMenagePiece (nil, nil, TOBEtudes) ;
  for Indice := 0 to TOBEtudes.detail.count - 1 do
  begin
    TOBLOC := TOBEtudes.detail [Indice] ;
    TOBLOC.deletedb (true) ;
  end;
end;

end.
