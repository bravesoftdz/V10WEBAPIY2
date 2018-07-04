unit ULibExercice;

interface

Uses
{$IFNDEF EAGLCLIENT}
     DB,
{$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF TTW}
 uWa ,
{$ENDIF}
{$IFNDEF EAGLSERVER}
 controls,  // pout TControl
 stdctrls,  // pour TEdit
{$ENDIF}
 SysUtils, // pour le decode date
 Ent1,      // pour le VH
 HEnt1,     // pour V_PGI
 HMsgBox,   // pour le PGIInfo
 ParamSoc ,
 UtilPGI ,
 HCtrls,
 uTOB       // TOB
  {$IFDEF MODENT1}
  , CPTypeCons
  {$ENDIF MODENT1}
 ;


type

 TZExercice = Class
  private

   FTOBExo                  : TOB ;
   FTOBEnCours              : TOB ;
   FTOBSuivant              : TOB ;
   FTOBPrecedent            : TOB ;
   FTOBExoV8                : TOB ;
   FTOBExoEdit              : TOB ;
   FTOBEntree               : TOB ;
   FTOBCPExoRef             : TOB ;
   FTOBNMoins2              : TOB ;
   FDetailClot              : integer ;
   FDossier                 : string ;

   function GetEnCours      : TExoDate ;
   function GetSuivant      : TExoDate ;
   function GetPrecedent    : TExoDate ;
   function GetEntree       : TExoDate ;
   function GetExoV8        : TExoDate ;
   function GetExoEdit      : TExoDate ;
   function GetCPExoRef     : TExoDate ;
   function GetNMoins2      : TExoDate ;
   function GetExoClo       : TTabExo ;
   function GetExercices ( Index : integer )  : TExoDate ;
   function GetNbExercices  : Integer;

  public

   constructor Create( Parle : boolean = true ; vDossier : String = '' ) ; virtual ;
   destructor  Destroy ; override ;

   procedure   ChargeMagExo ;

   function EstExoClos  ( vStCodeExo : string ) : boolean ;
   function QuelDateExo ( vDtDateC : TDateTime ; var Date1,Date2 : TDateTime) : string ;
   function QuelExoDate ( vDtDate1, vDtDate2 : TDateTime ; var vMonoExo : Boolean ; var vExo : TExoDate) : Boolean ; overload ;
   function QuelExo     ( vDtDateC : string ; vRechEnBase : boolean = false ) : string;
   function QuelExoDT   ( vDtDateC : TDateTime ; vRechEnBase : boolean = false ) : String;
   function QuelExoDate ( vDtDateC : TDateTime ) : TExoDate ; overload ;
   function QuelExoDate ( vStCodeExo : string ) : TExoDate ; overload ;
   function local : boolean ;

   property EnCours         : TExoDate   read GetEnCours ;
   property Entree          : TExoDate   read GetEntree ;
   property CPExoRef        : TExoDate   read GetCPExoRef ;
   property Suivant         : TExoDate   read GetSuivant ;
   property Precedent       : TExoDate   read Getprecedent ;
   property ExoV8           : TExoDate   read GetExoV8 ;
   property ExoEdit         : TExoDate   read GetExoEdit ;
   property DetailClot      : integer    read FDetailClot ;
   property ExoClo          : TTabExo    read GetExoClo ;
   property NMoins2         : TExoDate   read GetNMoins2 ;
   property Dossier         : string     read FDossier ;
   property Exercices[Index : integer ]  : TExoDate read GetExercices ;
   property NbExercices     : Integer    read GetNbExercices;
 end ;

{$IFNDEF NOVH}
function  CExoRefOuvert ( vBoParle : boolean ) : Boolean;
function  CMajRequeteExercice ( ExoRelatif : string ; StSQL : string ) : string;
function  CRelatifVersExercice ( ExoRelatif : string ) : string;
function  CCalculExerciceRelatif ( iExo : integer ) : string;
function  CExerciceVersRelatif ( stCode : string ) : string;
{$IFNDEF EAGLSERVER}
procedure CExoRelatifToDates ( ExoRelatif : string ; ED1,ED2 : TControl; bFiltre : boolean = False ) ;
{$ENDIF} 
function  CGetExerciceNMoins2 : TExoDate;
{$ENDIF}
{$IFNDEF EAGLSERVER}
{$IFNDEF NOVH}
procedure CInitComboExercice ( CB : THValComboBox );
{$ENDIF}
{$ENDIF}
function  CQuelExercice( Date : TDateTime ; var Exo : TExoDate ) : boolean;
function  CDateDebutTrimestre(Date : TDateTime) : TDateTime;
function  CDateFinTrimestre(Date : TDateTime) : TDateTime;
function  CLibelleExerciceRelatif ( stCode, stLibelle : string ) : string;

function  CControleDureeExercice( vDateDebut, vDateFin: TDateTime) : Boolean;


// AJOUT ME 22-02-2005 changement d'unité de CPTEUTIL
function  WhatDate(St : String ; Var DD1,DD2 : tDateTime ; Var Err : Integer ; Var Exo : TExoDate ) : boolean ;
Function  QUELEXODATE(Date1,Date2 : TDateTime ; Var MonoExo : Boolean ; Var Exo : TExoDate ) : Boolean ;

function  CGetDaysPerMonth( vDate : string ; vExoVisu : TExoDate ) : integer ;


function  ctxExercice : TZExercice ;
procedure ReleaseCtxExercice ;

{$IFDEF MODENT1}
procedure NombrePerExo(Exo : TExoDate; var PremMois, PremAnnee, NombreMois : Word);
procedure RempliExoDate (var Exo : TExoDate ) ;
function  QuelDateDeExo(CodeExo : String3; var Exo : TExoDate) : Boolean ;
function  QuelExoDtBud(DD : TDateTime) : string;
{$IFNDEF EAGLSERVER}
procedure ExoToEdDates(Exo : String3; ED1, ED2 : TControl); {Gestion du cas où Exo = '' sinon appel de ExoToDates}
procedure ExoToDates  (Exo : String3; ED1, ED2 : TControl);
{$ENDIF EAGLSERVER}
{$ENDIF MODENT1}



implementation

{$IFDEF EAGLSERVER}
{$IFDEF NOVH}
Uses
 ULibCpContexte ;
{$ENDIF}
{$ENDIF}

var
 _Exercice : TZExercice ;

function ctxExercice : TZExercice ;
begin
{$IFDEF EAGLSERVER}
{$IFDEF NOVH}
 result := TCPContexte.GetCurrent.Exercice ;
{$ENDIF} 
{$ELSE}
 if _Exercice = nil then
  _Exercice := TZExercice.Create ;
 result := _Exercice ;
{$ENDIF}

end ;

procedure ReleaseCtxExercice ;
begin
 FreeAndNil(_Exercice) ;
end ;



{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/07/2001
Modifié le ... : 05/07/2001
Description .. : - Recherche si l'exercice de référence est sélectioné
                 - Teste si il est encore ouvert
Mots clefs ... : Exercice Ouvert Référence
*****************************************************************}
{$IFNDEF NOVH}
function CExoRefOuvert ( vBoParle : boolean ) : Boolean;
var i : integer ;
begin
  if CtxPCl in V_PGI.PgiContexte then
    Result := VH^.CPExoRef.Code <> ''
  else
  begin
    Result := True;
    Exit;
  end;

 if not result and vBoParle then
  PgiInfo('Vous devez renseigner l''exercice de référence','Vérification d''exercice');

 for i:=1 to 5 do if VH^.CPExoRef.Code=VH^.ExoClo[i].Code then
  begin
   if vBoParle then PgiInfo('l''exercice de référence est clos','Vérification d''exercice');
   Result := False;
   Exit;
  end; // for
end;
{$ENDIF} 


{==============================================================================}
function CDateDebutTrimestre( Date : TDateTime ) : TDateTime;
var
 Annee,Mois,Jour : Word;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.CDateDebutTrimestre') ;
{$ENDIF}
  DecodeDate(Date,Annee,Mois,Jour);
  case Mois of
  1,2,3 : Result := EncodeDate(Annee,1,1);
  4,5,6 : Result := EncodeDate(Annee,4,1);
  7,8,9 : Result := EncodeDate(Annee,7,1);
  10,11,12: Result := EncodeDate(Annee,10,1);
  else Result := Date;
  end;
end;

{==============================================================================}
function CDateFinTrimestre ( Date : TDateTime ) : TDateTime;
var
 Annee,Mois,Jour : Word;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.CDateFinTrimestre') ;
{$ENDIF}
  DecodeDate(Date,Annee,Mois,Jour);
  case Mois of
  1,2,3 : Result := EncodeDate(Annee,3,31);
  4,5,6 : Result := EncodeDate(Annee,6,30);
  7,8,9 : Result := EncodeDate(Annee,9,30);
  10,11,12: Result := EncodeDate(Annee,12,31);
  else Result := Date;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 28/07/2003
Modifié le ... :   /  /
Description .. : Renvoie l'exercice associé à la date passée en
Suite ........ : paramètre.
Mots clefs ... :
*****************************************************************}
{ DONE -olaurent -cdll serveur :  }
function CQuelExercice( Date : TDateTime ; var Exo : TExoDate) : boolean;
{$IFNDEF NOVH}
var i    : integer;
{$ENDIF}
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.CQuelExercice') ;
{$ENDIF}
{$IFNDEF NOVH}
  i := 1 ;  Result := False;
  while (VH^.Exercices[i].Code<>'') do
  begin
    if (Date >= VH^.Exercices[i].Deb) and  (Date <= VH^.Exercices[i].Fin) then
    begin
      Exo := VH^.Exercices[i];
      Result := true;
      break;
    end;
    Inc (i,1);
  end;
{$ELSE}
  Exo    := TCPContexte.GetCurrent.Exercice.QuelExoDate(Date) ;
  result := Exo.Code <> '' ;
{$ENDIF}
end;


{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/07/2003
Modifié le ... :   /  /
Description .. : Renseigne la liste des exercices relatifs dans une liste de
Suite ........ : choix
Suite ........ : Exercice relatif : N+1, N+2, N-1,N-2 ....
Suite ........ : Pour l'en-cours, on aura N+0
Mots clefs ... : 
*****************************************************************}
{$IFNDEF EAGLSERVER}
{$IFNDEF NOVH}
procedure CInitComboExercice ( CB : THValComboBox );

    function LibelleExercice ( DateDebut, DateFin : TDateTime) : string;
    var yd, md, dd : Word;
        yf, mf, df : Word;
    begin
      DecodeDate ( DateDebut, yd, md, dd );
      DecodeDate ( DateFin, yf, mf, df);
      if yd = yf then Result := Format('%.04d',[yf])
      else Result := Format('%.04d/%.04d',[yd,yf]);
    end;

var i       : integer;
    stCode, stLibelle  : string;
    TExoLib, T : TOB;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.CInitComboExercice') ;
{$ENDIF}
  if CB = nil then exit;
  { Suppression des éléments de la combo }
  CB.Items.Clear;
  CB.Values.Clear;
  if CB.Vide then
  begin
    if CB.VideString = '' then
      CB.Items.Add(TraduireMemoire('<<Tous>>'))
    else CB.Items.Add(TraduireMemoire(CB.VideString));
    CB.Values.Add('');
  end;
  { Chargement des libellés des exercices }
  TExoLib := TOB.Create('', nil, -1);
  try
    TExoLib.LoadDetailDB('EXERCICE','','',nil,False);
    { Constitution de la liste }
    i:=1;
    while (VH^.Exercices[i].Code<>'') do
    begin
      stCode := CCalculExerciceRelatif ( i );
      CB.Values.Add( stCode );
      T := TExoLib.FindFirst(['EX_EXERCICE'],[VH^.Exercices[i].Code],False);
      if T <> nil then stLibelle := LibelleExercice ( VH^.Exercices[i].Deb,VH^.Exercices[i].Fin)
      else stLibelle := '';
      CB.Items.Add(CLibelleExerciceRelatif ( stCode , stLibelle) );
      Inc (i,1);
    end;
  finally
    TExoLib.Free;
  end;
end;
{$ENDIF}
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/07/2003
Modifié le ... :   /  /
Description .. : Calcul du code de l'exercice relatif en fonction de son
Suite ........ : positionnement dans la chronologie des exercices
Mots clefs ... :
*****************************************************************}
{$IFNDEF NOVH}
function CCalculExerciceRelatif ( iExo : integer ) : string;
var i, iEnCours : integer;
begin
  Result := '';
  { Recherche de l'encours }
  iEnCours := -1;
  i := 1;
  while ((VH^.Exercices[i].Code<>'') and (iEnCours=-1)) do
  begin
    if VH^.Exercices[i].Code=VH^.Encours.Code then iEnCours := i
    else Inc (i,1);
  end;
  { En cours non trouvé, on sort }
  if iEnCours = - 1 then exit;
  { Calcul du code exercice relatif }
  if iExo < iEnCours then Result := 'N-'+IntToStr(iEnCours-iExo)
  else Result := 'N+'+IntToStr(iExo-iEnCours);
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/07/2003
Modifié le ... :   /  /
Description .. : Renvoie le libellé de l'exercice relatif ( différent du libellé réel
Suite ........ : de l'exercice)
Mots clefs ... :
*****************************************************************}
function CLibelleExerciceRelatif ( stCode, stLibelle : string ) : string;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.CLibelleExerciceRelatif') ;
{$ENDIF}
  if stCode = 'N-1' then Result :=  TraduireMemoire('Précédent')+' : '+stLibelle
  else if stCode = 'N+0' then Result :=  TraduireMemoire('En cours')+' : '+stLibelle
  else if stCode = 'N+1' then Result :=  TraduireMemoire('Suivant')+' : '+stLibelle
  else Result := stCode+' : '+stLibelle;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/07/2003
Modifié le ... :   /  /
Description .. : Récupère le code exercice réel en fonction du code
Suite ........ : exercice relatif.
Mots clefs ... :
*****************************************************************}
{$IFNDEF NOVH}
function CRelatifVersExercice ( ExoRelatif : string ) : string;
var i : integer;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.CRelatifVersExercice') ;
{$ENDIF}
  Result := '';
  if ExoRelatif='' then exit;
  if ExoRelatif[1]<>'N' then exit;
  i := 1;
  while (VH^.Exercices[i].Code<>'') do
  begin
    if CCalculExerciceRelatif ( i ) = ExoRelatif then
    begin
      Result := VH^.Exercices[i].Code;
      break;
    end;
    Inc(i,1);
  end;
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/07/2003
Modifié le ... :   /  /
Description .. : Renvoie le code relatif en fonction du code exercice réel
Mots clefs ... :
*****************************************************************}
{$IFNDEF NOVH}
function CExerciceVersRelatif ( stCode : string ) : string;
var i : integer;
    stRelatif : string;
begin
  stRelatif := '';
  i := 1;
  while (VH^.Exercices[i].Code<>'') do
  begin
    if VH^.Exercices[i].Code=stCode then
    begin
      stRelatif := CCalculExerciceRelatif ( i );
      break;
    end;
    Inc (i,1);
  end;
  Result := stRelatif;
end;
{$ENDIF} 

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/07/2003
Modifié le ... :   /  /
Description .. : Mise à jour du critère exercice relatif vers le code exercice
Suite ........ : réel dans une requête SQL
Mots clefs ... :
****************************************************************}
{$IFNDEF NOVH}
function CMajRequeteExercice ( ExoRelatif : string ; StSQL : string ) : string;
var stCode : string;
    stSQLTrad : string;
begin
  stSQLTrad := stSQL ;
  if ExoRelatif <> '' then
  begin
    stCode := CRelatifVersExercice ( ExoRelatif );
    if stCode <> '' then
    begin
      stSQLTrad := FindEtReplace(stSQLTrad,'EXERCICE = "'+ExoRelatif+'"','EXERCICE="'+stCode+'"',True);
      stSQLTrad := FindEtReplace(stSQLTrad,'EXERCICE = '''+ExoRelatif+'''','EXERCICE='''+stCode+'''',True);
    end;
  end;
  Result := stSQLTrad;
end;
{$ENDIF} 


{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/07/2003
Modifié le ... : 31/10/2003
Description .. : Mise à jour des date en fonction d'un code exercice relatif
Suite ........ : 31/10/2003 - CA - Gestion du cas des filtres
Mots clefs ... :
*****************************************************************}
{$IFNDEF NOVH}
{$IFNDEF EAGLSERVER}
procedure CExoRelatifToDates ( ExoRelatif : string ; ED1,ED2 : TControl; bFiltre : boolean = False ) ;
var stCode : string;
    DateDeb, DateFin, DateTmpDeb,DateTmpFin : TDateTime;
    Exo : TExoDate;
    yd, md, dd : word;
    ye, me, de : word;
begin
  if not bFiltre then
  begin
    stCode := CRelatifVersExercice ( ExoRelatif );
    if stCode = '' then exit;
    ExoToDates( stCode, ED1, ED2);
  end
  else
  begin
    stCode := CRelatifVersExercice ( ExoRelatif );
    if StCode = '' then Exit; // GCO - 20/07/2005 - FQ 15568  

    if IsValidDate ( TEdit(ED1).Text ) then
    begin
      DateDeb := StrToDate ( TEdit(ED1).Text );
      if IsValidDate ( TEdit(ED2).Text ) then
      begin
        DateFin := StrToDate ( TEdit(ED2).Text );
        Exo.Code:= CRelatifVersExercice ( ExoRelatif );
        RempliExoDate ( Exo );
        { Mise à jour de la date de début }
        DecodeDate( DateDeb, yd, md, dd );
        DecodeDate( Exo.Deb, ye, me, de );
        DateTmpDeb := EncodeDate (ye, md, dd );
        { Mise à jour de la date de fin }
        DecodeDate( DateFin, yd, md, dd );
        DecodeDate( Exo.Fin, ye, me, de );
        DateTmpFin := EncodeDate (ye, md, dd );
        if ((DateTmpDeb >=Exo.Deb) and (DateTmpDeb<=Exo.Fin)) and
              ((DateTmpFin >=Exo.Deb) and (DateTmpFin<=Exo.Fin)) then
        begin
          TEdit(ED1).Text := DateToStr( DateTmpDeb );
          TEdit(ED2).Text := DateToStr( DateTmpFin );
        end else
        begin
          TEdit(ED1).Text := DateToStr( Exo.Deb );
          TEdit(ED2).Text := DateToStr( Exo.Fin );
        end;
      end;
    end;
  end;
end;
{$ENDIF}
{$ENDIF} 

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/06/2004
Modifié le ... :   /  /
Description .. : Contrôle le nombre de mois entre les bornes d'un exercice
Suite ........ : afin de vérifier son authenticité dans PGI ( Max 24 mois )
Mots clefs ... :
*****************************************************************}
function CControleDureeExercice( vDateDebut, vDateFin: TDateTime) : Boolean;
var lPremMois, lPremAnnee, lNombreMois : Word;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.CControleDureeExercice') ;
{$ENDIF}
  NombreMois( vDateDebut, vDateFin, lPremMois, lPremAnnee, lNombreMois );
  Result := ( lNombreMois <= 24 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/02/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFNDEF NOVH}
function CGetExerciceNMoins2 : TExoDate;
var i : integer;
begin
  FillChar( Result, SizeOf(Result), #0);

  // Si pas d'exercice précédent, pas besoin de continuer
  if VH^.Precedent.Code = '' then Exit;

  i := 1;
  while (VH^.Exercices[i].Code <> '') do
  begin
    if (VH^.Exercices[i].Code = VH^.Precedent.Code) and (i > 1) then
    begin
      Result := VH^.Exercices[i-1];
      Break;
    end;
    Inc (i,1);
  end;

end;
{$ENDIF}
////////////////////////////////////////////////////////////////////////////////


function _DecodeDateOle ( var St: String ; var SD : String) : Boolean;
var
 Pos1,Pos2 : Integer ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice._DecodeDateOle') ;
{$ENDIF}
 SD        := '' ;
 Result    := false ;
 Pos1      := Pos('(',St) ;
 Pos2      := Pos(')',St) ;
 if ( Pos1 > 0 ) and ( Pos2 > 0 ) Then
  begin
   sd          := Copy(St,Pos1+1,Pos2-Pos1-1) ;
   St[Pos1]    := ' ' ;
   St[Pos2]    := ' ' ;
   Result      := true ;
  end ;
end ;

function _TraiteDate( St : String ; Var DD : TDateTime) : Integer ;
var
 i        : integer ;
 St1,StD  : string ;
 JJ,MM,AA : word ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice._TraiteDate') ;
{$ENDIF}

 result := -13 ;
 DD     := 0 ;

 if ( length(St) <> 8 ) and ( length(St) <> 10 ) then exit ;

 for i := 1 to length(St) do
  begin
   if ( St[i] in ['0'..'9','-','/'] ) = false then exit ;
   if ( St[i] = '-' ) or ( St[i] = '/' ) then St[i] := ';' ;
  end ;

 St      := St + ';' ;
 St1     := ReadTokenSt(St) ;
 JJ      :=StrToInt(St1) ;
 AA      :=0 ;
 MM      :=0 ;

 if St <> '' then
  begin
   St1 := ReadTokenSt(St) ;
   MM  := StrToInt(St1) ;
  end ;

 if St <> '' Then
  begin
   St1    := ReadTokenSt(St) ;
   AA     := StrToInt(St1) ;
   if AA < 80 then
    AA := 2000 + AA
     else
      if  AA <= 99 then
       AA := 1900 + AA ;
  end ;

 if ( JJ = 0 ) or ( MM = 0 ) or ( AA = 0) then Exit ;
 if JJ in [1..31] = false then Exit ;
 if MM in [1..12] = false then Exit ;

 StD    := FormatFloat('00',JJ) + '/' + FormatFloat('00',MM) + '/' + FormatFloat('00',AA) ;
 if not IsValidDate(StD) Then Exit ;
 
 DD     := EncodeDate( AA , MM , JJ) ;
 Result :=0 ;
 
end ;


procedure GetSemaine( St : string ; var DD1,DD2 : TDateTime ; var Err : Integer ; var Exo : TExoDate ) ;
var
  i,j,NumSemaine,NumSemaine2,Annee,PremierJourAnnee : integer ;
  DDT : TDateTime ;
  lMono : boolean ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.GetSemaine') ;
{$ENDIF}

 Err       := -13 ;
 DD1       := StrToDate(StDate1900) ;
 DD2       := StrToDate(StDate2099) ;
 Exo.code  := '' ;
 i         := pos( ':' , St ) ;

 if i = 0 then exit ;

 System.delete(St,1,i) ;
 i := pos(':',St) ;
 j := pos('-',St) ;
 if ( i = 0 ) or ( j = 0 ) then exit ;

try
 NumSemaine  := StrToInt( copy(St,1,j-1) ) ;
 NumSemaine2 := StrToInt( copy(St,j+1,i-j-1) ) ;
 System.delete(St,1,i) ;
 Annee       := StrToInt(St) ;
except
 exit ;
end ;

 if ( Annee < 0 ) or ( Annee > 9999 ) then exit ;
 if ( NumSemaine < 1 ) or ( NumSemaine2 < 1) then exit  ;
 if NumSemaine > NumSemaine2 then NumSemaine2 := NumSemaine2 + 52 ;
 if Annee < 100 then if annee > 80 then Annee := Annee + 1900 else Annee := Annee +2000 ;

 PremierJourAnnee := DayOfWeek ( encodedate(Annee,1,1)) ;
 if PremierJourAnnee > 2 then
  DDT := EncodeDate( Annee, 1, 10-PremierJourAnnee)
   else
    if PremierJourAnnee = 1 then
     DDT := EncodeDate(Annee,1,2)
      else
       DDT := EncodeDate(Annee,1,1) ;

 DD1   := DDT + ( NumSemaine - 1 ) * 7 ;
 DD2   := DDT + ( NumSemaine2 - 1) * 7 + 6 ;

 QuelExoDate(DD1,DD2,lMono,Exo) ;
 if not lMono then Exo.Code := '' ;

end ;

Function NbMoins(St : String) : Integer ;
Var i,l : Integer ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.NbMoins') ;
{$ENDIF}
l:=0 ; Repeat i:=Pos('-',St) ; If i>0 Then BEGIN St[i]:=' ' ; Inc(l) ; END ; Until i=0 ;
Result:=-1*l ;
end ;


Function NbPer(St : String) : Integer ;
Var l,i : Integer ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.NbPer') ;
{$ENDIF}
l:=Length(St) ; Result:=0 ;
For i:=1 To L Do
  BEGIN
  If St[i] in ['0'..'9'] Then
     BEGIN
     If (i<L) And (St[i+1] In ['0'..'9']) Then Result:=StrToInt(Copy(St,i,2))
                                          Else Result:=StrToInt(Copy(St,i,1)) ;
     Exit ;
     END ;
  END ;
end ;

Function TraitePeriodique(NoPer,Coeff : Integer ; Exo : TExoDate ; Var DD1,DD2 : TdateTime ) : Integer ;
Var DD,D1,D2 : TdateTime ;
    ii : Integer ;
    OkOk : Boolean ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.TraitePeriodique') ;
{$ENDIF}
Result:=0 ;
If NoPer<>0 Then
   BEGIN
   DD1:=Exo.Deb ;
   If NoPer>1 Then DD1:=PlusMois(Exo.Deb,Coeff*(NoPer-1)) ;
   If DD1>Exo.Fin Then BEGIN Result:=-15 ; Exit ; END Else
      BEGIN
      DD2:=PlusMois(DD1,Coeff)-1 ;
      If DD2>Exo.Fin Then BEGIN Result:=-16 ; Exit ; END ;
      END ;
   END Else
   BEGIN
   DD:=V_PGI.DateEntree ; D1:=Exo.Deb ; D2:=PlusMois(DD1,Coeff)-1 ; ii:=0 ; //OkOk:=FALSE ;
   Repeat
     Inc(ii) ;
     OkOk:=(DD>=D1) And (DD<=D2) ;
     If Not OkOk Then
        BEGIN
        D1:=D2+1 ; D2:=PlusMois(D1,Coeff)-1 ;
        END ;
   Until (ii>20) Or OkOk ;
   If OkOk Then
      BEGIN
      DD1:=D1 ; DD2:=D2 ;
      If DD2>Exo.Fin Then BEGIN Result:=-16 ; Exit ; END ;
      END Else Result:=-15 ;
   END ;
end ;


function WhatDate(St : String ; Var DD1,DD2 : tDateTime ; Var Err : Integer ; Var Exo : TExoDate ) : boolean;
Var OnExo,OnPeriode,OnTrimestre,OnSemestre,OnBimestre,OnQuatromestre,OnJour : Boolean ;
    DepuisDebut,JusquaFin : Boolean ;
    NoPer : Integer ;
    Decal : Integer ;
    a,m,j,aa,mm,jj : Word ;
    MonoExo : Boolean ;
    Sd1,Sd2 : String ;
    DD3 : tDateTime ;
    {$IFDEF NOVH}
    lExo : TZExercice ;
    ii,l : Integer ;
    {$ENDIF}
{ Semaine :
W;2;2000 ou W:2-3:00-> Semaine 2 an 2000
W;13;1999 ou w;13;99 -> Semaine 13 an 1999
Renvoie : DD1,DD2 : les 2 dates de la semaine
          Exo : tExo date correspondant à la semaine
}
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.WhatDate') ;
{$ENDIF}
 {$IFDEF EAGLSERVER}
 {$IFDEF NOVH}
  lExo := TCPContexte.GetCurrent.Exercice ;
 {$ENDIF}
 {$ENDIF}
 
 Err             := 0 ;
 Exo.Code        := '' ;
 OnExo           := FALSE ;
 OnPeriode       := FALSE ;
 OnTrimestre     := FALSE ;
 OnSemestre      := FALSE ;
 OnBimestre      := FALSE ;
 OnQuatromestre  := FALSE ;
 DepuisDebut     := FALSE ;
 JusquaFin       := FALSE ;
 NoPer           := 0 ;
 Decal           := 0 ;
 OnJour          := FALSE ;
 DD1             := GetEnCours.Deb ;
 DD2             := GetEnCours.Fin ;

 if _DecodeDateOLE(St,SD1) and _DecodeDateOLE(St,SD2) Then
  begin
   Err := _TraiteDate(Sd1,DD1) + _TraiteDate(Sd2,DD2) ;
   if Err = 0 then
    begin
     if  DD1 > DD2 then
      begin
       DD3 := DD1 ;
       DD1 := DD2 ;
       DD2 := DD3 ;
      end;
      { CA - 01/06/2006 - je supprime ces 2 lignes, je ne sais pas pourquoi elles sont arrivées là !!! }
//       else
//        Err := - 13 ;
     QuelExoDate(DD1,DD2,MonoExo,Exo) ;
     if ( not MonoExo) then Err := -15 ;
     if ( Exo.Code = '') then Err := -6 ;
    END ;
   result := Err = 0 ;
   exit ;
  end ;

 if ( Pos('N',St) > 0 ) or ( pos('n',St) > 0 ) then OnExo := true else
  if ( pos('M',St) > 0 ) or ( pos('m',St) > 0 ) then OnPeriode := true else
   if ( pos('T',St) > 0 ) or ( Pos('t',St) > 0 ) then OnTrimestre := true else
    if ( pos('B',St) > 0 ) or ( pos('b',St) > 0 ) then OnBimestre := true else
     if ( Pos('S',St) > 0 ) or ( pos('s',St) > 0 ) then OnSemestre := true else
      if ( Pos('Q',St) > 0 ) or ( Pos('q',St) > 0 ) then OnQuatromestre := true else
       if ( Pos('J',St) > 0 ) or ( Pos('j',St) > 0 ) then OnJour := true else
        if ( Pos('W',St) > 0 ) or ( Pos('w',St) > 0 ) then
         begin
           GetSemaine(St,DD1,DD2,Err,Exo) ;
           result := Err = 0 ;
           exit ;
         end
          else
           Err := -1 ;
If Err=0 Then
   BEGIN
   If (Pos('<',St)>0) Then DepuisDebut:=TRUE Else
      If (Pos('>',St)>0) Then JusquaFin:=TRUE ;
   If Pos('+',St)>0 Then Decal:=+1 Else If Pos('-',St)>0 Then Decal:=NbMoins(St) ;
   NoPer:=NbPer(St) ;
   END ;
If (Not (OnPeriode or OnJour)) And (DepuisDebut Or JusquaFin) Then Err:=-13 ;
{$IFNDEF NOVH}
If Decal=1 Then Exo:=VH^.Suivant ;
If Decal=0 Then Exo:=VH^.EnCours ;
{$ENDIF}
{$IFDEF EAGLSERVER}
{$IFDEF NOVH}
If Decal=1 Then Exo:=lExo.Suivant ;
If Decal=0 Then Exo:=lExo.EnCours ;
{$ENDIF}
{$ENDIF}
If Decal<0 Then
   BEGIN
   {$IFDEF NOVH}
   l:=0 ;
   For ii:=5 DownTo 1 Do
     BEGIN
     If lExo.ExoClo[ii].Code<>'' Then Inc(l) ;
     If l=Abs(Decal) Then Exo:=lExo.ExoClo[ii] ;
(*
     {$IFNDEF NOVH}
     If VH^.ExoClo[ii].Code<>'' Then Inc(l) ;
     If l=Abs(Decal) Then Exo:=VH^.ExoClo[ii] ;
     {$ENDIF}
*)
     END ;
    {$ELSE}
    // CA - 26/04/2007 pour problème ExoClo limité à 5
    QuelDateDeExo(CRelatifVersExercice ( 'N'+IntToStr(Decal)),Exo);
    {$ENDIF}
   END ;
If OnJour Then
   BEGIN
    QuelExoDate(V_PGI.DateEntree,V_PGI.DateEntree,MonoExo,Exo) ;
   END ;
If Exo.Code=''  Then Err:=-6 ;
If Err=0 Then
   BEGIN
   If OnPeriode Then
      BEGIN
      If NoPer>0 Then
         BEGIN
         DD1:=Exo.Deb ;
         { FQ 15656 - CA - 22/04/2005 : pas le premier mois de l'exercice,
          on se base sur la date de début de mois (cas des exercices qui ne commencent pas au 01 }
         //If NoPer>1 Then DD1:=PlusMois(Exo.Deb,NoPer-1) ;
         If NoPer>1 Then DD1:=PlusMois(DebutDeMois(Exo.Deb),NoPer-1) ;
         If DepuisDebut Then BEGIN DD2:=FinDeMois(DD1) ; DD1:=Exo.Deb ; END Else
            If JusquaFin Then DD2:=Exo.Fin Else
               BEGIN
               DD2:=FinDeMois(DD1) ;
               END ;
         If DD1>Exo.Fin Then Err:=-17 ;
         If DD2>Exo.Fin Then Err:=-18 ;
         END Else
         BEGIN
         DD1:=V_PGI.DateEntree ;
         DecodeDate(DD1,a,m,j) ;
         DecodeDate(Exo.Deb,aa,mm,jj) ;
         DD1:=EncodeDate(aa,m,1) ;
         DD2:=FinDeMois(DD1) ;
         If DepuisDebut Then DD1:=Exo.Deb Else If JusquaFin Then DD2:=Exo.Fin ;
         END ;
      END ;
   If OnExo Then
      BEGIN
      DD1:=Exo.Deb ; DD2:=Exo.Fin ;
      END ;
   If OnTrimestre Then Err:=TraitePeriodique(NoPer,3,Exo,DD1,DD2) ;
   If OnSemestre Then Err:=TraitePeriodique(NoPer,6,Exo,DD1,DD2) ;
   If OnBimestre Then Err:=TraitePeriodique(NoPer,2,Exo,DD1,DD2) ;
   If OnQuatromestre Then Err:=TraitePeriodique(NoPer,4,Exo,DD1,DD2) ;
   If OnJour Then
      BEGIN
      DD1:=V_PGI.DateEntree ; DD2:=V_PGI.DateEntree ;
      If DepuisDebut Then DD1:= Exo.Deb Else If JusquaFin Then DD2:=Exo.Fin ;
      END ;
   END ;

 result := Err = 0 ;

end ;


Function DansExo(Exo : TExoDate ; Date1,Date2 : TDateTime) : Boolean ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.DansExo') ;
{$ENDIF}
Result:=(Date1>=Exo.Deb) And (Date2<=Exo.Fin) ;
end ;


(*======================================================================*)
Function DansExoCloture(DD1,DD2 : TDateTime ; Var lExo : TExoDate) : Boolean ;
Var
 {$IFNDEF NOVH}
 i : Integer ;
 {$ENDIF}
{$IFDEF EAGLSERVER}
 lBoMono : boolean ;
{$ENDIF}
BEGIN
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.DansExoCloture') ;
{$ENDIF}
{ DONE -olaurent -cdll serveur :  }
{$IFNDEF NOVH}
Result:=FALSE ;
For i:=1 To 5 Do
  BEGIN
   If DansExo(VH^.ExoClo[i],DD1,DD2) then BEGIN LExo:=VH^.ExoClo[i] ; Result:=TRUE ; Exit ; END ;
  END ;
{$ELSE}
{$IFDEF EAGLSERVER}
 TCPContexte.GetCurrent.Exercice.QuelExoDate(DD1,DD2,lBoMono,lExo) ;
 result := TCPContexte.GetCurrent.Exercice.EstExoClos(lExo.Code) ;
 {$ENDIF}
{$ENDIF}
END ;

Function DansExoViaQuery(Date1,Date2 : TDateTime ; Var MonoExo : Boolean ; Var Exo1,Exo2 : TExoDate) : Boolean ;
Var Q : TQuery ;
    Deb,Fin : TDateTime ;
    PremFois : Boolean ;
    Exo : TExoDate ;
    i,j,k : Word ;
BEGIN
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.DansExoViaQuery') ;
{$ENDIF}
Result:=FALSE ;
Q:=OpenSQL('SELECT * FROM EXERCICE',TRUE,-1,'',true) ; PremFois:=TRUE ;
While Not Q.Eof Do
   BEGIN
   Deb:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
   Fin:=Q.FindField('EX_DATEFIN').AsDateTime ;
   If ((date1>=Deb) And (date1<=Fin)) Or ((date2>=Deb) And (date2<=Fin))then
      BEGIN
      If PremFois Then MonoExo:=TRUE Else MonoExo:=FALSE ;
      Exo.Deb:=Deb ; Exo.Fin:=Fin ; Exo.Code:=Q.FindField('EX_EXERCICE').AsString ;
      Exo.DateButoir:=Q.FindField('EX_DATECUM').AsDateTime ;
      Exo.DateButoirRub:=Q.FindField('EX_DATECUMRUB').AsDateTime ;
      Exo.DateButoirBud:=Q.FindField('EX_DATECUMBUD').AsDateTime ;
      Exo.DateButoirBudgete:=Q.FindField('EX_DATECUMBUDGET').AsDateTime ;
      NombrePerExo(Exo,i,j,k) ; Exo.NombrePeriode:=k ;
      If PremFois Then Exo1:=Exo Else Exo2:=Exo ;
      PremFois:=FALSE ;
      Result:=TRUE ;
      END ;
   Q.Next ;
   END ;
Ferme(Q) ;

END ;


constructor TZExercice.Create ( Parle : boolean = true ; vDossier : String = ''  ) ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.Create') ;
{$ENDIF}
 inherited Create ;

 FDossier := vDossier ;

 FTOBExo := nil ;
 ChargeMagExo ;
end ;


destructor TZExercice.Destroy;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.Destroy') ;
{$ENDIF}
 FreeAndNil(FTOBExo) ;
 inherited;
end;

function CTOBVersExo( vTOB : TOB ) : TExoDate;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.CTOBVersExo') ;
{$ENDIF}
 FillChar(result,SizeOf(result),#0) ;

 try

 if vTOB = nil then exit ;
 result.Code              := vTOB.GetValue('EX_EXERCICE') ;
 result.Deb               := vTOB.GetValue('EX_DATEDEBUT') ;
 result.Fin               := vTOB.GetValue('EX_DATEFIN') ;
 result.DateButoir        := vTOB.GetValue('EX_DATECUM') ;
 result.DateButoirRub     := vTOB.GetValue('EX_DATECUMRUB') ;
 result.DateButoirBud     := vTOB.GetValue('EX_DATECUMBUD') ;
 result.DateButoirBudgete := vTOB.GetValue('EX_DATECUMBUDGET') ;
 result.EtatCpta          := vTOB.GetValue('EX_ETATCPTA') ; // AJOUT ME
 if isNumeric(vTOB.GetValue('PER')) then
  result.NombrePeriode     := vTOB.GetValue('PER') ; // AJOUT ME

  except
  on E : Exception do
   begin
    {$IFNDEF EAGLSERVER}
     PGIError('ULibExercice.CTOBVersExo : ' + E.Message ) ;
    {$ENDIF}
    {$IFDEF TTW}
    cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.CTOBVersExo : ' + E.Message ) ;
    {$ENDIF}
   end ;
 end ;

end;


function TZExercice.EstExoClos( vStCodeExo : string ) : boolean;
var
 i : integer ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.EstExoClos') ;
{$ENDIF}
 result := false ;
 for i := 0  to FDetailClot - 1 do
  begin
   result := vStCodeExo = FTOBExo.Detail[i].GetValue('EX_EXERCICE') ;
   if result then break ;
  end ;// while
end;

function TZExercice.GetCPExoRef: TExoDate;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.GetCPExoRef') ;
{$ENDIF}
 result := CTOBVersExo( FTOBCPExoRef ) ;
end;

function TZExercice.GetEnCours : TExoDate;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.GetEnCours') ;
{$ENDIF}
 Result := CTOBVersExo( FTOBEnCours ) ;
end;

function TZExercice.GetEntree : TExoDate;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.GetEntree') ;
{$ENDIF}
 result := CTOBVersExo( FTOBEntree ) ;
end;

function TZExercice.GetExoEdit : TExoDate;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.GetExoEdit') ;
{$ENDIF}
 result := CTOBVersExo( FTOBExoEdit ) ;
end;

function TZExercice.GetExoV8 : TExoDate;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.GetExoV8') ;
{$ENDIF}
 result := CTOBVersExo( FTOBExoV8 ) ;
end;

function TZExercice.GetPrecedent : TExoDate;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.GetPrecedent') ;
{$ENDIF}
 result := CTOBVersExo( FTOBPrecedent  ) ;
 {$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.GetPrecedent finis') ;
{$ENDIF}
end;

function TZExercice.GetSuivant : TExoDate;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.GetSuivant') ;
{$ENDIF}
 result := CTOBVersExo( FTOBSuivant ) ;
end;

function TZExercice.GetNMoins2 : TExoDate ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.GetNMoins2') ;
{$ENDIF}
result := CTOBVersExo( FTOBNMoins2 ) ;
end ;

function TZExercice.GetExoClo : TTabExo ;
var
 i : integer ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.GetExoClo') ;
{$ENDIF}
 for i := 1 to 5 do with result[i] do begin Deb := 0 ; Fin := 0 ; Code := '' ; end ;
 for i := 0  to FDetailClot - 1 do
  result[i] := CTOBVersExo(FTOBExo.Detail[i]) ;
end ;

function TZExercice.QuelDateExo ( vDtDateC : TDateTime ; var Date1,Date2 : TDateTime) : string ;
var
 lExo : TExoDate ;
 i    : Integer ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.QuelDateExo') ;
{$ENDIF}

 for i := 0 to FTOBExo.Detail.Count - 1 do
  begin
   lExo := CTOBVersExo(FTOBExo.Detail[i]) ;
   if (vDtDateC >= lExo.Deb) and ( vDtDateC <= lExo.Fin) then
    begin
     Result := lExo.Code ;
     Date1  := lExo.Deb ;
     Date2  := lExo.Fin ;
     break ;
    end ;
  end ; // for

end ;

function TZExercice.QuelExo ( vDtDateC : string ; vRechEnBase : boolean = false ) : string ;
var
 lQ : TQuery ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.QuelExo') ;
{$ENDIF}
 Result := QuelExoDT ( StrToDate( vDtDateC ) ) ;
 if ( Result = '' ) and vRechEnBase then
  begin
  lQ := OpenSQL('SELECT * FROM EXERCICE WHERE EX_DATEDEBUT<="' + USDATETIME(StrToDate(vDtDateC)) + '" AND EX_DATEFIN>="' + USDATETIME(StrToDate(vDtDateC)) + '" ' , true, - 1 , '' , true , FDossier) ;
  if not lQ.EOF then result := lQ.FindField('EX_EXERCICE').AsString ;
  Ferme(lQ) ;
 end ; // if
end ;


function TZExercice.QuelExoDT ( vDtDateC : TDateTime ; vRechEnBase : boolean = false ) : String ;
var
 lQ : TQuery ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.QuelExoDT') ;
{$ENDIF}
 result := QuelExoDate(vDtDateC).Code ;
 if ( Result = '' ) and vRechEnBase then
  begin
  lQ := OpenSQL('SELECT * FROM EXERCICE WHERE EX_DATEDEBUT<="' + USDATETIME(vDtDateC) + '" AND EX_DATEFIN>="' + USDATETIME(vDtDateC) + '" ' , true, - 1 , '' , true , FDossier) ;
  if not lQ.EOF then result := lQ.FindField('EX_EXERCICE').AsString ;
  Ferme(lQ) ;
 end ; // if
end ;

function TZExercice.QuelExoDate ( vDtDateC : TDateTime ) : TExoDate ;
var
 lExo : TExoDate ;
 i    : Integer ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.QuelExoDate(vDtDateC ') ;
{$ENDIF}

 FillChar( Result, SizeOf(Result), #0) ;

 for i := 0 to FTOBExo.Detail.Count - 1 do
  begin
   lExo := CTOBVersExo(FTOBExo.Detail[i]) ;
   if (vDtDateC >= lExo.Deb) and ( vDtDateC <= lExo.Fin) then
    begin
     Result := lExo ;
     break ;
    end ;
  end ; // for

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 26/05/2005
Modifié le ... :   /  /    
Description .. : Retourne le TExoDate associé au code exercice passé en 
Suite ........ : paramètre
Mots clefs ... : 
*****************************************************************}
function TZExercice.QuelExoDate ( vStCodeExo : string ) : TExoDate ;
var
 lExo : TExoDate ;
 i    : Integer ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.QuelExoDate(vDtDateC ') ;
{$ENDIF}

 FillChar( Result, SizeOf(Result), #0) ;

 for i := 0 to FTOBExo.Detail.Count - 1 do
  begin
   lExo := CTOBVersExo(FTOBExo.Detail[i]) ;
   if (vStCodeExo = lExo.Code) then
    begin
     Result := lExo ;
     break ;
    end ;
  end ; // for
end ;

function _DansExo( Exo : TExoDate ; Date1,Date2 : TDateTime) : Boolean ;
begin
 Result:=(Date1>=Exo.Deb) And (Date2<=Exo.Fin) ;
end ;

function TZExercice.QuelExoDate( vDtDate1, vDtDate2 : TDateTime ; var vMonoExo : Boolean ; var vExo : TExoDate) : Boolean ;
var
 lExo : TExoDate ;
 i    : integer ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.QuelExoDate (vDtDate1, vDtDate2 ') ;
{$ENDIF}

 vMonoExo := false ;
 result   := false ;

 for i := 0 to FTOBExo.Detail.Count - 1 do
  begin
   lExo := CTOBVersExo(FTOBExo.Detail[i]) ;

   if _DansExo(lExo , vDtDate1 , vDtDate2 ) then
    begin
     result  := true ;
     vMonoExo := true ;
     vExo     := lExo ;
     exit ;
    end ; // if
  end ; // for

end ;

procedure TZExercice.ChargeMagExo ;
var
 lQ           : TQuery ;
 lTOB         : TOB ;
 lBoPremier   : boolean ;
 i            : integer ;
 lStExo       : string ;

 procedure _Affecte( var vTOBSource,vTOBDest : TOB ) ;
 var
  PremMois     : Word ;
  PremAnnee    : Word ;
  NombreMois   : Word ;
  begin
   vTOBDest := vTOBSource ;
   if vTOBDest <> nil then
    begin
     NombrePerExo(CTOBVersExo(vTOBDest), PremMois, PremAnnee, NombreMois) ;
     vTOBDest.PutValue('PER', NombreMois ) ;
    end ; // if
  end ;

begin
{$IFDEF TTW}
 cWA.MessagesAuClient('IMPORT','','ULibExercice.ChargeMagExo') ;
{$ENDIF}

 FTOBEnCours     := nil ;
 FTOBSuivant     := nil ;
 FTOBPrecedent   := nil ;
 FTOBExoV8       := nil ;
 FTOBExoEdit     := nil ;
 FTOBEntree      := nil ;
 lBoPremier      := true ;
 FDetailClot     := 0 ;

 try

 if assigned(FTOBExo) then
  FTOBExo.Free ;
 FTOBExo         := TOB.Create('',nil,-1) ;

 lQ := OpenSQL('SELECT * FROM EXERCICE ORDER BY EX_DATEDEBUT', true, - 1 , 'EXERCICE' , true , FDossier) ;


 FTOBExo.LoadDetailDB('EXERCICE','','',lQ,true) ;
 if FTOBExo.Detail.Count > 0 then
  FTOBExo.Detail[0].AddChampSupValeur('PER',0,true) ;
 Ferme(lQ) ;

 if local
   then lStExo := GetParamSocSecur('SO_EXOV8','' )
   else lStExo := GetParamSocDossierSecur('SO_EXOV8','' , FDossier) ;

 for i := 0 to FTOBExo.Detail.Count - 1 do
  begin

   lTOB := FTOBExo.Detail[i] ;

   if lTOB.GetValue('EX_EXERCICE') = Trim( lStExo ) then
    _Affecte(lTOB,FTOBExoV8) ;

   if lTOB.GetValue('EX_ETATCPTA') = 'CDE' then
    begin
     _Affecte(lTOB,FTOBPrecedent) ;

     if ( i > 0 ) then
      FTOBNMoins2 := FTOBExo.Detail[i-1] ;

     Inc(FDetailClot) ;
    end // if CDE
     else
      if ( lTOB.GetValue('EX_ETATCPTA') = 'OUV' ) or ( lTOB.GetValue('EX_ETATCPTA') = 'CPR' ) then
       begin
        if lBoPremier then
         begin
          lBoPremier  := false ;
          _Affecte(lTOB,FTOBEnCours) ;
         end // if
          else
           begin
            _Affecte(lTOB,FTOBSuivant) ;
            break ;
           end ;
       end ;

  end ; // for

 if ( V_PGI.DateEntree >= EnCours.Deb ) and ( V_PGI.DateEntree <= Encours.Fin ) then
  _Affecte(FTOBEncours,FTOBEntree)
   else
    if (V_PGI.DateEntree >= Suivant.Deb) and (V_PGI.DateEntree <= Suivant.Fin) then
     _Affecte(FTOBSuivant,FTOBEntree) ;
     
 if Entree.Code='' then
  _Affecte(FTOBEncours,FTOBEntree) ;

 if local
   then lStExo := GetParamSocSecur('SO_CPEXOREF','' )
   else lStExo := GetParamSocDossierSecur('SO_CPEXOREF','' , FDossier) ;

  FTOBCPExoRef := FTOBExo.FindFirst(['EX_EXERCICE'],[lStExo],true) ;
 _Affecte(FTOBCPExoRef,FTOBCPExoRef) ;

 except
  on E : Exception do
   begin
    {$IFNDEF EAGLSERVER}
     PGIError('ULibExercice.ChargeMagExo : ' + E.Message ) ;
    {$ENDIF}
    {$IFDEF TTW}
    cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.ChargeMagExo : ' + E.Message ) ;
    {$ENDIF}
   end ;
 end ;

end;

function TZExercice.GetExercices ( Index : integer ) : TExoDate ;
begin
 result := CTOBVersExo( FTOBExo.Detail[Index] ) ;
end ;

function TZExercice.GetNbExercices  : Integer;
begin
  Result := FTOBExo.Detail.Count;
end;

{$IFDEF EAGLSERVER}
{$IFDEF NOVH}
Function QuelExoDate(Date1,Date2 : TDateTime ; Var MonoExo : Boolean ; Var Exo : TExoDate ) : Boolean ;
Var LExo,LExo2 : TExoDate ;
lZExo : TZExercice ;
BEGIN
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.Function QuelExoDate(Date1,Date2... ') ;
{$ENDIF}
lZExo := TCPContexte.GetCurrent.Exercice ;
MonoExo:=FALSE ; Result:=FALSE ;
if DansExo(lZExo.Precedent,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo:=lZExo.Precedent ; END else
   if DansExo(lZExo.EnCours,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo:=lZExo.EnCours ; END else
      if DansExo(lZExo.Suivant,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo:=lZExo.Suivant ; END Else
         If DansExoCloture(Date1,Date2,LExo) Then BEGIN MonoExo:=TRUE ; Exo:=LExo ; END Else
            If DansExoViaQuery(Date1,Date2,MonoExo,LExo,LExo2) Then Exo:=LExo Else Exit ;
Result:=TRUE ;
END ;
{$ENDIF}
{$ENDIF}

{$IFNDEF NOVH}
Function QuelExoDate(Date1,Date2 : TDateTime ; Var MonoExo : Boolean ; Var Exo : TExoDate) : Boolean ;
Var LExo,LExo2 : TExoDate ;
BEGIN
MonoExo:=FALSE ; Result:=FALSE ;
if (VH^.Precedent.Code<>'') and DansExo(VH^.Precedent,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo:=VH^.Precedent ; END else
   if (VH^.enCours.Code<>'') and DansExo(VH^.EnCours,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo:=VH^.EnCours ; END else
      if (VH^.Suivant.Code<>'') and DansExo(VH^.Suivant,Date1,Date2) then BEGIN MonoExo:=TRUE ; Exo:=VH^.Suivant ; END Else
         If DansExoCloture(Date1,Date2,LExo) Then BEGIN MonoExo:=TRUE ; Exo:=LExo ; END Else
            If DansExoViaQuery(Date1,Date2,MonoExo,LExo,LExo2) Then Exo:=LExo Else Exit ;
Result:=TRUE ;
END ;
{$ENDIF}


function CGetDaysPerMonth( vDate : string ; vExoVisu : TExoDate ) : integer ;
var
 lYear,lDay,lMonth : word ;
begin
 DecodeDate(StrToDate(vDate), lYear, lMonth, lDay) ;
 result:=DaysPerMonth(lYear, lMonth) ;
 if EncodeDate(lYear,lMonth,result) >= vExoVisu.Fin then
  begin
   DecodeDate(vExoVisu.Fin, lYear, lMonth, lDay) ;
   result := lDay ;
  end ;
end ;

function TZExercice.local: boolean;
begin
  result := (FDossier = '') or ( FDossier = V_PGI.SchemaName ) ;
end;

{$IFDEF MODENT1}
{---------------------------------------------------------------------------------------}
procedure NombrePerExo(Exo : TExoDate; var PremMois, PremAnnee, NombreMois : Word);
{---------------------------------------------------------------------------------------}
var
  PremJour,
  DernAnnee,
  DernMois,
  DernJour : Word ;
begin
  if Exo.Deb > Exo.Fin then
    NombreMois := 0
  else begin
    DecodeDate(Exo.Deb, PremAnnee, PremMois, PremJour);
    DecodeDate(Exo.Fin, DernAnnee, DernMois, DernJour);
    NombreMois := 12 * (DernAnnee - PremAnnee) + (DernMois - PremMois + 1);
  end;
end;

{$IFNDEF EAGLSERVER}
{JP 14/08/07 : FQ 20091 : gestion du cas où Exo est vide. Ce n'est pas fait dans ExoToDates, car
               personne n'est trop sûr sur d'éventuels impacts
{---------------------------------------------------------------------------------------}
procedure ExoToEdDates(Exo : String3; ED1, ED2 : TControl);
{---------------------------------------------------------------------------------------}
begin
  if Exo = '' then begin
    TEdit(ED1).Text := StDate1900;
    TEdit(ED2).Text := StDate2099;
  end
  else
    ExoToDates(Exo, ED1, ED2);
end;

{---------------------------------------------------------------------------------------}
procedure ExoToDates ( Exo : String3 ; ED1,ED2 : TControl ) ;
{---------------------------------------------------------------------------------------}
var
  D1,D2 : TDateTime ;
  Q     : TQuery;
  Okok  : boolean ;
begin
  {JP 14/08/07 : FQ 20091 : Si l'on veut gérer le cas où Exo vaut '', appeler ExoToEdDates définie ci-dessus}
  if EXO = '' then Exit;

  Okok := True;
  D1 := Date;
  D2 := Date;

  if EXO = VH^.Precedent.Code then begin
    D1 := VH^.Precedent.Deb;
    D2 := VH^.Precedent.Fin;
  end
  else if EXO = VH^.EnCours.Code then begin
    D1 := VH^.Encours.Deb;
    D2 := VH^.Encours.Fin;
  end
  else if EXO = VH^.Suivant.Code then begin
    D1 := VH^.Suivant.Deb;
    D2 := VH^.Suivant.Fin;
  end
  else begin
    Q := OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE = "' + Exo + '"', True,-1,'',true);
    if not Q.EOF then begin
      D1 := Q.FindField('EX_DATEDEBUT').asDateTime;
      D2 := Q.FindField('EX_DATEFIN').asDateTime;
    end else
      Okok:=False ;
    Ferme(Q) ;
  end;

  if Okok then begin
    TEdit(ED1).Text := DateToStr(D1);
    TEdit(ED2).Text := DateToStr(D2);
  end;
end;
{$ENDIF EAGLSERVER}

{---------------------------------------------------------------------------------------}
procedure RempliExoDate (var Exo : TExoDate ) ;
{---------------------------------------------------------------------------------------}
begin
  if Exo.Code = GetPrecedent.Code then begin
    Exo.Deb := GetPrecedent.Deb;
    Exo.Fin := GetPrecedent.Fin;
    Exo.NombrePeriode := GetPrecedent.NombrePeriode;
  end
  else if Exo.Code=GetEnCours.Code then begin
      Exo.Deb := GetEnCours.Deb;
      Exo.Fin := GetEnCours.Fin;
      Exo.NombrePeriode := GetEnCours.NombrePeriode;
  end
  else if Exo.Code = GetSuivant.Code then begin
    Exo.Deb := GetSuivant.Deb;
    Exo.Fin := GetSuivant.Fin;
    Exo.NombrePeriode := GetSuivant.NombrePeriode;
  end
  else
    {Ajout ME 02/10/2003 pour les exercice n-2}
    QuelDateDeExo(Exo.Code, Exo);
end;

{---------------------------------------------------------------------------------------}
Function QuelDateDeExo(CodeExo : String3 ; Var Exo : TExoDate) : Boolean ;
{---------------------------------------------------------------------------------------}
Var i : Integer ;
    Q : TQuery ;
BEGIN
{$IFDEF NOVH}
Exo:=GetEnCours ; Result:=FALSE ;
if CodeExo=GetPrecedent.Code then BEGIN Result:=TRUE ; Exo:=GetPrecedent ; END else
   if CodeExo=GetEnCours.Code then BEGIN Result:=TRUE ; Exo:=GetEnCours ; END else
      if CodeExo=GetSuivant.Code then BEGIN Result:=TRUE ; Exo:=GetSuivant ; END Else
         BEGIN
         For i:=1 To 5 Do BEGIN If CodeExo=GetExoClo[i].Code then
            BEGIN Exo:=GetExoClo[i] ; Result:=TRUE ; Exit ; END ; END ;
          // ajout me
              if TCPContexte.GetCurrent.Exercice.Exercices[1].Code <> '' then
              begin
                    for  i:= 1  to TCPContexte.GetCurrent.Exercice.NbExercices do
                    begin
                                if (TCPContexte.GetCurrent.Exercice.Exercices[i].Code = CodeExo) then
                                begin
                                     Exo.Code    := TCPContexte.GetCurrent.Exercice.Exercices[i].Code ;
                                     Exo.Deb     := TCPContexte.GetCurrent.Exercice.Exercices[i].Deb ;
                                     Exo.Fin     := TCPContexte.GetCurrent.Exercice.Exercices[i].Fin ;
                                     Result      := TRUE ;
                                     break;
                                end;
                    end;
              end
              else
              begin
                   Q:=OpenSQL('SELECT * FROM EXERCICE WHERE EX_EXERCICE="'+CodeExo+'"',TRUE,-1,'',true) ;
                   If Not Q.Eof Then
                      BEGIN
                      Exo.Code:=Q.FindField('EX_EXERCICE').AsString ;
                      Exo.Deb:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
                      Exo.Fin:=Q.FindField('EX_DATEFIN').AsDateTime ;
                      Result:=TRUE ;
                      END ;
                   Ferme(Q) ;
              end;
        END ;
{$ELSE}
Exo:=VH^.EnCours ; Result:=FALSE ;
if CodeExo=VH^.Precedent.Code then BEGIN Result:=TRUE ; Exo:=VH^.Precedent ; END else
   if CodeExo=VH^.EnCours.Code then BEGIN Result:=TRUE ; Exo:=VH^.EnCours ; END else
      if CodeExo=VH^.Suivant.Code then BEGIN Result:=TRUE ; Exo:=VH^.Suivant ; END Else
         BEGIN
         For i:=1 To 5 Do BEGIN If CodeExo=VH^.ExoClo[i].Code then
            BEGIN Exo:=VH^.ExoClo[i] ; Result:=TRUE ; Exit ; END ; END ;
          // ajout me
              if VH^.Exercices[1].Code <> '' then
              begin
                    for  i:= 1  to length (VH^.Exercices) do
                    begin
                                if (VH^.Exercices[i].Code = CodeExo) then
                                begin
                                     Exo.Code    := VH^.Exercices[i].Code ;
                                     Exo.Deb     := VH^.Exercices[i].Deb ;
                                     Exo.Fin     := VH^.Exercices[i].Fin ;
                                     Result      := TRUE ;
                                     break;
                                end;
                    end;
              end
              else
              begin
                   Q:=OpenSQL('SELECT * FROM EXERCICE WHERE EX_EXERCICE="'+CodeExo+'"',TRUE,-1,'',true) ;
                   If Not Q.Eof Then
                      BEGIN
                      Exo.Code:=Q.FindField('EX_EXERCICE').AsString ;
                      Exo.Deb:=Q.FindField('EX_DATEDEBUT').AsDateTime ;
                      Exo.Fin:=Q.FindField('EX_DATEFIN').AsDateTime ;
                      Result:=TRUE ;
                      END ;
                   Ferme(Q) ;
              end;
        END ;
{$ENDIF NOVH}
END ;

{---------------------------------------------------------------------------------------}
function QuelExoDtBud(DD : TDateTime) : string;
{---------------------------------------------------------------------------------------}
Var i : Integer ;
    Q : Tquery ;
begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','ULibExercice.QuelExoDtBUD') ;
{$ENDIF}
Result := GetEnCours.Code ;
If (dd>= GetEnCours.Deb) and (dd <= GetEnCours.Fin) then Result := GetEnCours.Code else
   If (dd >= GetSuivant.Deb) and (dd <= GetSuivant.Fin) then Result := GetSuivant.Code Else
      If (dd >= GetPrecedent.Deb) and (dd <= GetPrecedent.Fin) then Result := GetPrecedent.Code Else
      BEGIN
      For i:=1 To 5 Do
        BEGIN
        If (dd >= GetExoClo[i].Deb) And (dd <= GetExoClo[i].Fin)then BEGIN Result := GetExoClo[i].Code ; Exit ; END ;
        END ;
      Q:=OpenSQL('SELECT * FROM EXERCICE WHERE EX_DATEDEBUT<="'+USDATETIME(DD)+'" AND EX_DATEFIN>="'+USDATETIME(DD)+'" ' ,TRUE,-1,'',true) ;
      if Not Q.EOF then Result:=Q.FindField('EX_EXERCICE').AsString ;
      Ferme(Q) ;
      END ;
end ;
{$ENDIF MODENT1}




initialization
finalization
 ReleaseCtxExercice ;


end.




