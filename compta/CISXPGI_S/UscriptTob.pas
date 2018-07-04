unit UScriptTob;


interface
uses Classes, Windows, SysUtils, Forms, StdCtrls,
  Dialogs, FileCtrl, inifiles, menus,
  HEnt1, HMsgBox, Hctrls, UScript, 
{$IFDEF CISXPGI}
  uYFILESTD,
  UtilPGI,
  Math,
  HRichOle,
  HRichEdt,
  ULibWindows,
{$ENDIF}
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet, ADODB,{$ENDIF}
{$ENDIF}
  UTOB
  , cbpPath; // TcbpPath.GetCegidUserTempPath

{$IFNDEF CISXPGI}
Type LaVariablesCISX = RECORD
     Ini, Domaine, Mode,
     Nature, Complement,
     Directory, Script,
     NomFichier, Option,
     Famille, Appelant,
     RepSTD, LISTECOMMANDE,
     Monotraitement, compress : string;
     ATob                     : TOB;
     ListeFichier             : string;
    end;
Var VHCX : ^LaVariablesCISX ;
FUNCTION  CHARGECIX (FichierIni : string=''): Boolean ;
procedure LibereCX;
function GetInfoVHCX : LaVariablesCISX;

{$ELSE}

{$IFDEF EAGLCLIENT}
type TZVarCisx = class
  private
    Ini,
    LISTECOMMANDE    : string;
  public
    Domaine                  : string;
    NomFichier               : string;
    Directory                : string;
    Appelant                 : string;
    Famille                  : string;
    Mode                     : string;
    Script                   : string;
    Nature                   : string;
    Complement               : string;
    Option                   : string;
    ATob                     : TOB;
    Monotraitement, compress : string;
    ListeFichier             : string;
    constructor Create ; virtual ;
    procedure CHARGECIX(FichierIni : string='');
    destructor  Destroy ; override ;

end;
function GetInfoVHCX : TZVarCisx;
function InitScriptAuto(table : string) : integer;
{$ENDIF}

function RendTobparametre  (Domaine : string; var TOBParam : TOB; crit2 : string='COMPTA'; crit3: string='PARAM'; Crit4 : string=''): integer;
function Recupprefixe( tt : string) : string;
Function ReturnReleveMax (TableTr,CpteGeneral : string) : integer;
Function MaxRelevEtebac(TableTr,NumCpte : string) : integer;
function ReturnComptebancaire(ETABBQ,GUICHET,Numerodecompte : string) : string;
procedure ChargementListeTable (var ListBox1 : TListBox; Domaine : string);
procedure ChargementPays (var TobDomaine: TOB);
{$ENDIF}

implementation
{$IFDEF CISXPGI}
 uses ULibCpContexte;
{$ENDIF}

{$IFNDEF CISXPGI}
FUNCTION  CHARGECIX (FichierIni : string=''): Boolean ;
var
Commande, tmp       : string;
sTempo              : string;
FicIni              :TIniFile;
CurrentDir          : string;
ListeFile,Domaine   : string;
TFV                 : TOB;
begin
  Result := TRUE; FicIni := nil;
     if ParamCount > 0 then  // si lancement par la ligne de commande
     begin
          if FichierIni = '' then
          begin
             Commande := ParamStr(1);
             tmp := ReadTokenPipe (Commande, ';');
             CurrentDir := ReadTokenPipe (Commande, ';');
             if CurrentDir = '' then CurrentDir := ParamStr(1);
          end
          else
             currentDir := FichierIni;

          if (Copy(currentDir, 1, 5) = '/INI=') then
          begin
               sTempo := Copy(currentDir, 6, length(CurrentDir));
               if not FileExists(sTempo) then
               begin
                    PGIInfo ('Le fichier '+ sTempo+ ' n''existe pas','');
                    FicIni.free;
                    Result := FALSE;
                    exit;
               end;
               FicIni        := TIniFile.Create(sTempo);
               New(VHCX) ; FillChar(VHCX^,Sizeof(VHCX^),#0) ;
               VHCX^.Ini        := sTempo;
               VHCX^.Domaine    := FicIni.ReadString ('PRODUIT', 'Domaine', '');
               VHCX^.Mode       := FicIni.ReadString ('PRODUIT', 'Mode', '');
               VHCX^.Nature     := FicIni.ReadString ('PRODUIT', 'NATURE', '');
               VHCX^.Complement := FicIni.ReadString ('PRODUIT', 'COMPLEMENT', '');
               VHCX^.Option     := FicIni.ReadString ('PRODUIT', 'OPTION', '');
               VHCX^.Famille    := FicIni.ReadString ('PRODUIT', 'FAMILLE', '');
               VHCX^.Appelant   := FicIni.ReadString ('PRODUIT', 'APPELANT', '');
               //VHCX^.RepSTD     := FicIni.ReadString ('PRODUIT', 'REPSTD', '');
               CurrentDossier     := FicIni.ReadString ('PRODUIT', 'REPSTD', '');
               CurrentDonnee   := FicIni.ReadString ('PRODUIT', 'REPDOS', '');
               if not DirectoryExists(CurrentDonnee) then CreateDir(CurrentDonnee);

               VHCX^.LISTECOMMANDE  := FicIni.ReadString ('PRODUIT', 'LISTECOMMANDE', '');
               VHCX^.ATob := nil;
               if VHCX^.LISTECOMMANDE <> '' then
               begin
                     ListeFile := VHCX^.LISTECOMMANDE;
                     Commande  := ListeFile;
                     Domaine   := VHCX^.Domaine;
                     While Commande <> '' do
                     begin
                          Commande    := ReadTokenPipe(ListeFile, ';');
                          if Commande <> '' then
                          begin
                               if VHCX^.ATob = nil then VHCX^.ATob := TOB.Create('COMMANDE', nil,-1);
                               TFV := TOB.Create(Commande, VHCX^.ATob, -1);
                               TFV.AddChampSupValeur('REPERTOIRE', FicIni.ReadString (Commande, 'REPERTOIRE', ''));
                               TFV.AddChampSupValeur('SCRIPT', FicIni.ReadString (Commande, 'SCRIPT', ''));
                               TFV.AddChampSupValeur('NOMFICHIER', FicIni.ReadString (Commande, 'NOMFICHIER', ''));
                               TFV.AddChampSupValeur('LISTEFICHIER', FicIni.ReadString (Commande, 'LISTEFICHIER', ''));
                               Domaine := ReadTokenPipe(VHCX^.Domaine, ';');
                               TFV.AddChampSupValeur('DOMAINE', Domaine);
                          end;
                     end;
               end
               else
               begin
                              VHCX^.Directory  := FicIni.ReadString ('COMMANDE', 'REPERTOIRE', '');
                              VHCX^.Script     := FicIni.ReadString ('COMMANDE', 'SCRIPT', '');
                              VHCX^.NomFichier := FicIni.ReadString ('COMMANDE', 'NOMFICHIER', '');
                              VHCX^.ListeFichier:= FicIni.ReadString ('COMMANDE', 'LISTEFICHIER', '');
                              VHCX^.Monotraitement := FicIni.ReadString ('COMMANDE', 'MONOTRAITEMENT', '');
                              VHCX^.Compress       := FicIni.ReadString ('COMMANDE', 'COMPRESS', '');
               end;
               FicIni.free;
          end;
     end;
end;

procedure LibereCX;
begin
     if (ParamCount > 0) and assigned(VHCX) then
     begin
          if VHCX^.ATOB <> nil then
          begin VHCX^.ATOB.free; VHCX^.ATOB := nil end;
          Dispose(VHCX) ;
     end;
end;
{$ENDIF}
{$IFDEF CISXPGI}
function RendTobparametre  (Domaine : string; var TOBParam : TOB; crit2 : string='COMPTA'; crit3: string='PARAM'; Crit4 : string=''): integer;
var
Filename                   : string;
CodeRetour                 : integer;
entete                     : boolean;
encode                     : string;

begin
        Filename := GetWindowsTempPath +'PGI\STD\CISX\';
        if Domaine <> '' then Filename := Filename + Domaine+'\';
        if crit2 <> '' then Filename := Filename + crit2 +'\';
        if crit3 <> '' then Filename := Filename + crit3 +'\';
        if crit4 <> '' then Filename := Filename +crit4 +'\';
        Filename := Filename +  V_PGI.LanguePrinc+'\CEG\'+
        Domaine+'Compta.Cix';
        if FileExists (Filename) then
                      DeleteFile(Filename);

(* CA - 05/07/2007 - Remplacement DosPath car n'existe pas en EAGL

        if V_PGI.DosPath[length (V_PGI.DosPath)] <> '\' then
        Filename :=  V_PGI.DosPath +'\'+Domaine+'Compta.Cix'
        else Filename :=  V_PGI.DosPath +Domaine+'Compta.Cix';

*)
  Filename :=  TcbpPath.GetCegidUserTempPath +Domaine+'Compta.Cix';


        CodeRetour :=  AGL_YFILESTD_EXTRACT (Filename, 'CISX', ExtractFileName(Filename), Domaine, Crit2, Crit3, crit4, '', FALSE, V_PGI.LanguePrinc, 'CEG');
        if CodeRetour = -1 then
        begin
              if TOBParam <> nil then TOBParam.free;
              TOBParam := TOB.create ('', nil, -1);
              TOBParam.LoadFromXMLFile(Filename, entete, encode);
        end;
        Result := CodeRetour;
end;

{$IFDEF EAGLCLIENT}
procedure TZVarCisx.CHARGECIX(FichierIni : string='');
var
Commande, tmp       : string;
sTempo              : string;
FicIni              :TIniFile;
CurrentDir          : string;
ListeFile           : string;
TFV                 : TOB;
begin
     FicIni := nil;
     if (FichierIni <> '') or (pos ('/INI=', ParamStr(1)) <> 0)then  // si lancement par la ligne de commande
     begin
          if FichierIni = '' then
          begin
             Commande := ParamStr(1);
             tmp := ReadTokenPipe (Commande, ';');
             CurrentDir := ReadTokenPipe (Commande, ';');
             if CurrentDir = '' then CurrentDir := ParamStr(1);
          end
          else
             currentDir := FichierIni;

          if (Copy(currentDir, 1, 5) = '/INI=') then
          begin
               sTempo := Copy(currentDir, 6, length(CurrentDir));
               if not FileExists(sTempo) then
               begin
                    PGIInfo ('Le fichier '+ sTempo+ ' n''existe pas','');
                    FicIni.free;
                    exit;
               end;
               FicIni        := TIniFile.Create(sTempo);
               Ini        := sTempo;
               Domaine    := FicIni.ReadString ('PRODUIT', 'Domaine', '');
               Mode       := FicIni.ReadString ('PRODUIT', 'Mode', '');
               Nature     := FicIni.ReadString ('PRODUIT', 'NATURE', '');
               Complement := FicIni.ReadString ('PRODUIT', 'COMPLEMENT', '');
               Option     := FicIni.ReadString ('PRODUIT', 'OPTION', '');
               Famille    := FicIni.ReadString ('PRODUIT', 'FAMILLE', '');
               Appelant   := FicIni.ReadString ('PRODUIT', 'APPELANT', '');
               CurrentDossier     := FicIni.ReadString ('PRODUIT', 'REPSTD', '');
               CurrentDonnee   := FicIni.ReadString ('PRODUIT', 'REPDOS', '');
               if not DirectoryExists(CurrentDonnee) then CreateDir(CurrentDonnee);

               LISTECOMMANDE  := FicIni.ReadString ('PRODUIT', 'LISTECOMMANDE', '');
               if LISTECOMMANDE <> '' then
               begin
                     ListeFile := LISTECOMMANDE;
                     Commande  := ListeFile;
                     While Commande <> '' do
                     begin
                          Commande    := ReadTokenPipe(ListeFile, ';');
                          if Commande <> '' then
                          begin
                               if ATob = nil then ATob := TOB.Create('COMMANDE', nil,-1);
                               TFV := TOB.Create(Commande, ATob, -1);
                               TFV.AddChampSupValeur('REPERTOIRE', FicIni.ReadString (Commande, 'REPERTOIRE', ''));
                               TFV.AddChampSupValeur('SCRIPT', FicIni.ReadString (Commande, 'SCRIPT', ''));
                               TFV.AddChampSupValeur('NOMFICHIER', FicIni.ReadString (Commande, 'NOMFICHIER', ''));
                               TFV.AddChampSupValeur('LISTEFICHIER', FicIni.ReadString (Commande, 'LISTEFICHIER', ''));
                               TFV.AddChampSupValeur('MONOTRAITEMENT', FicIni.ReadString (Commande, 'MONOTRAITEMENT', ''));
                               TFV.AddChampSupValeur('COMPRESS', FicIni.ReadString (Commande, 'COMPRESS', ''));
                               Domaine := ReadTokenPipe(Domaine, ';');
                               TFV.AddChampSupValeur('DOMAINE', Domaine);
                          end;
                     end;
               end
               else
               begin
                              Directory  := FicIni.ReadString ('COMMANDE', 'REPERTOIRE', '');
                              Script     := FicIni.ReadString ('COMMANDE', 'SCRIPT', '');
                              NomFichier := FicIni.ReadString ('COMMANDE', 'NOMFICHIER', '');
                              ListeFichier:= FicIni.ReadString ('COMMANDE', 'LISTEFICHIER', '');
                              Monotraitement := FicIni.ReadString ('COMMANDE', 'MONOTRAITEMENT', '');
                              Compress       := FicIni.ReadString ('COMMANDE', 'COMPRESS', '');
               end;
               FicIni.free;
          end;
     end
     else
                            Domaine    := 'X';
end;

constructor TZVarCisx.Create ;
begin
     ATob := TOB.Create('COMMANDE', nil,-1);
     inherited ;
end;

destructor TZVarCisx.Destroy;
begin
  if atob <> nil then
  FreeAndNil(ATob);
  inherited ;
end;

function InitScriptAuto(table : string) : integer;
var
S             : TmemoryStream;
AStreamTable  : TmemoryStream;
Ascript       : TScript;
isc           : integer;
St            : string;
TF1,TOBVar    : TOB;
QP            : TQuery;
begin
S := nil; St := ''; AStreamTable := nil; TOBVar:= nil;
Result := -1; //Script inexistant
try
  QP := OpenSQl ('SELECT * FROM CPGZIMPREQ Where CIS_CLE="'+table+'"', TRUE);

  AStreamTable := TmemoryStream.create;
  if not QP.EOF then
  begin
       Result := 0; // Pas de variables
       s := TmemoryStream.create;
       TBlobField(QP.FindField('CIS_PARAMETRES')).SaveToStream (s);
       s.Seek (0,0);

       AScript := LoadScriptFromStream(s, AStreamTable);
       if AScript.Variables <> nil then
       begin
            Result := 1; // Script et variables trouvés
            TOBVar := TOB.Create('Enregistrement', nil, -1);
            For isc:=0 To AScript.Variables.count-1 do
            begin
                       St := AScript.Variables.Items[isc].Libelle;
                       St := ReadTokenpipe (St,'=');
                       TF1 := TOB.Create ('',TOBVar,-1);
                       TF1.AddChampSupValeur('Name',AScript.Variables.Items[isc].Name);
                       TF1.AddChampSupValeur('Libelle',St);
                       TF1.AddChampSupValeur('Text', AScript.Variables.Items[isc].Text);
                       TF1.AddChampSupValeur('Demandable', AScript.Variables.Items[isc].demandable);
            end;
            (* CA - 05/07/2007 - Remplacement DosPath car n'existe pas en EAGL

            if FileExists(V_PGI.DosPath + '\Echanges\Echanges_LSR.INI') then
               SysUtils.DeleteFile (V_PGI.DosPath + '\Echanges\Echanges_LSR.INI');
            TOBVar.SaveToFile(V_PGI.DosPath + '\Echanges\Echanges_LSR.INI',True,True,True);

            *)
            if FileExists(TcbpPath.GetCegidUserTempPath + 'Echanges\Echanges_LSR.INI') then
               SysUtils.DeleteFile (TcbpPath.GetCegidUserTempPath + 'Echanges\Echanges_LSR.INI');
            TOBVar.SaveToFile(TcbpPath.GetCegidUserTempPath + 'Echanges\Echanges_LSR.INI',True,True,True);
       end;
  end
  Finally
    Ferme (QP);
    s.free;
    AStreamTable.free;
    TOBVar.free;
  end;
end;

{$ENDIF}


{$IFNDEF CISXPGI}
function GetInfoVHCX : LaVariablesCISX;
{$ELSE}
function GetInfoVHCX : TZVarCisx;
{$ENDIF}
begin
{$IFNDEF CISXPGI}
 if VHCX = nil then exit;
 Result := VHCX^;
{$ELSE}
 Result := TCPContexte.GetCurrent.VarCisx;
{$ENDIF}
end ;
{$ENDIF}
function Recupprefixe( tt : string) : string;
var
Q  : Tquery;
begin
              Q := OpenSQL ('SELECT DT_PREFIXE from DETABLES WHERE DT_NOMTABLE="'+tt+'"', TRUE);
              Result := Q.FindField ('DT_PREFIXE').asstring;
              ferme (Q);
end;

Function ReturnReleveMax (TableTr,CpteGeneral : string) : integer;
var
Num, NumReleve : integer;
QQ             : TQuery;
begin
        QQ:=OpenSQL('SELECT max(EE_NUMERO) FROM ' + GetTableDossier('', TableTr) + ' WHERE EE_GENERAL = "'+CpteGeneral+'"',true) ;
        NumReleve := QQ.Fields[0].AsInteger;
        Ferme(QQ) ;
        QQ:=OpenSQL('SELECT max(CEL_NUMRELEVE) FROM ' + GetTableDossier('', 'EEXBQLIG') + ' WHERE CEL_GENERAL = "'+CpteGeneral+'"',true) ;
        Num := maxintvalue([NumReleve,QQ.Fields[0].AsInteger]);
        Ferme(QQ) ;
        if Num > NumReleve then NumReleve := Num;
        inc(NumReleve);
        Result := NumReleve;
end;

Function MaxRelevEtebac(TableTr,NumCpte : string) : integer;
var
QQ        : TQuery;
NumReleve : integer;
begin
    QQ:=OpenSQL('SELECT max(CET_NUMRELEVE) FROM ' + GetTableDossier('', TableTr) + ' WHERE CET_NUMEROCOMPTE = "'+NumCpte+'"',true) ;
    NumReleve := QQ.Fields[0].AsInteger;
    Ferme(QQ) ;
    inc(NumReleve);
    Result := NumReleve;
end;

function ReturnComptebancaire(ETABBQ,GUICHET,Numerodecompte : string) : string;
var
Q : TQuery;
begin
      Q := Opensql ('SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_ETABBQ="'+ETABBQ+'" AND '+
      'BQ_GUICHET="'+GUICHET+'" AND BQ_NUMEROCOMPTE="'+Numerodecompte+'"', TRUE);
      if not Q.EOF then
       Result :=  Q.FindField ('BQ_GENERAL').asstring;
      Ferme (Q);
end;

procedure ChargementListeTable (var ListBox1 : TListBox; Domaine : string);
var
QP    : TQuery;
begin
        if (Domaine = 'X') then
        begin
             QP := OpenSQL ('SELECT * FROM DETABLES WHERE DT_DOMAINE="C" AND DT_NOMTABLE  like "%EEXBQ%" or '+
             ' DT_NOMTABLE like "%ETEBAC%" or  DT_NOMTABLE="BANQUECP"', TRUE);

             while not QP.EOF do
             begin
                 if ListBox1.items.IndexOf(QP.FindField ('DT_NOMTABLE').asstring) = -1 then
                      ListBox1.Items.Add(QP.FindField ('DT_NOMTABLE').asstring);
                 QP.Next;
             end;
             Ferme(QP);
        end;
end;

procedure ChargementPays (var TobDomaine: TOB);
var
Q     : TQuery;
TobL  : TOB;
begin
Q := OpenSQL ('select * from PAYS', TRUE);
if TobDomaine = nil then TobDomaine := TOB.Create('', nil, -1);
While not Q.EOF do
begin
        TobL := TOB.Create('', TobDomaine, -1);
        TobL.AddChampSupValeur('Domaine', Q.FindField('PY_PAYS').asstring, False);
        TobL.AddChampSupValeur('Libelle', Q.FindField('PY_LIBELLE').asstring, False);
        Q.next;
end;
Ferme (Q);
end;

end.
