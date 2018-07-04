{***********UNITE*************************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 17/12/2002
Modifié le ... :   /  /
Description .. : Contient les fonctions qui permettent de gérer un server
Suite ........ : SQL (local). Ainsi que les déclarations de ces fonctions
Suite ........ : executablent depuis les scripts
Mots clefs ... : S1;DATABASE;
*****************************************************************}
unit uWABDMSACCESS;
{$IFDEF BASEEXT} azertyuiop {$ENDIF BASEEXT}

interface
uses
   utob, Classes, HenT1, MajTable, hctrls, uHttp, stdctrls, uWABD, uWAIni, uWAInfoIni
   {$IFNDEF eAGLCLient}
   ,
   uDbxDataSet, 
   hqry
   {$ENDIF eAGLClient}
    ;

Type
   cWABDMSACCESS = class(cWABD)
   private
   public
      constructor Create ; override ;
      destructor Destroy; override;
      {$IFNDEF EAGLCLIENT}
      function CreerBaseFrom(nomBase: string): boolean ;  overload ; override ;
      function SupprimerBase(): boolean ; override ;
      function ArchiverBase(): boolean ; override ;
      function RestorerBase(): boolean ; override ;
      function AttacherBase(): boolean ; override ;
      function DetacherBase(): boolean ; override ;
      function ExisteBase(): boolean ; override ;
      function TestConnexion(): boolean ; override ;
      {$ENDIF EAGLCLIENT}
      {$IFNDEF EAGLCLIENT}
      class function Action(UneAction: String; RequestTOB: TOB;var ResponseTOB: TOB): boolean ;
      {$ENDIF EAGLCLIENT}
   end ;

implementation

uses
   Windows, SysUtils, Forms, controls, (* FileCtrl, *)
   hmsgbox, M3fp, registry , inifiles, licutil
   {$IFNDEF BASEEXT}
   , ParamSoc
   {$ENDIF BASEEXT}
   ;


//var  cpte: integer = 0 ;

{$IFNDEF EAGLCLIENT}

function cWABDMSACCESS.CreerBaseFrom(nomBase: string): boolean ;
begin
result:=true ;
end ;

function cWABDMSACCESS.SupprimerBase(): boolean ;
begin
result:=true ;
end ;

function cWABDMSACCESS.ArchiverBase(): boolean ;
begin
result:=true ;
end ;

function cWABDMSACCESS.RestorerBase(): boolean ;
begin
result:=true ;
end ;

function cWABDMSACCESS.AttacherBase(): boolean ;
begin
result:=true ;
end ;

function cWABDMSACCESS.DetacherBase(): boolean ;
begin
result:=true ;
end ;

function cWABDMSACCESS.ExisteBase(): boolean ;
begin
result:=true ;
end ;

function cWABDMSACCESS.TestConnexion: boolean ;
begin
result:=true ;
end ;
{$ENDIF EAGLCLIENT}

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 06/12/2002
Modifié le ... : 06/12/2002
Description .. : Cronstructeur de l'objet "outil"
Mots clefs ... : S1;;DATABASE;
*****************************************************************}
constructor cWABDMSACCESS.Create ;
begin
inherited create() ;
ReaffecteNomTable(className,true) ;
driver:='' ;
end ;
{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 06/12/2002
Modifié le ... : 06/12/2002
Description .. : Destructeur de l'objet "outil"
Mots clefs ... : S1;;DATABASE;
*****************************************************************}
destructor cWABDMSACCESS.Destroy;
begin
inherited ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 23/01/2004
Modifié le ... : 23/01/2004
Description .. : Fonction cdgi de gestion de base 
Suite ........ :
Suite ........ : Attention :
Suite ........ :   format JOBDATE (yyymmdd)
Suite ........ :   format JOBTIME (hhmm00)
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLCLIENT}
class function cWABDMSACCESS.Action(UneAction: String; RequestTOB: TOB;var ResponseTOB: TOB): boolean ;
var
  T1: TOB ;
  ioWABDMSACCESS: cWABDMSACCESS ;
BEGIN
result:=true ; ioWABDMSACCESS:=nil ; T1:=nil ;
TraceExecution('dbssPGI.cWABDMSACCESS.'+UneAction+' :');
try
  ioWABDMSACCESS:=cWABDMSACCESS(cWABD(RequestTOB).clone) ;

  with ioWABDMSACCESS do
    begin
    try
           if UneAction='EXISTDB'           then result:=ExisteBase() ;
    except
      on E: Exception do
        ErrorMessage:=E.Message ;
      end ;
    setRetour(ResponseTOB,result,T1,nil,'') ;
    end ;
finally
  if T1<>nil then T1.free ;
  if ioWABDMSACCESS<>nil then ioWABDMSACCESS.free ;
  end ;
END ;
{$ENDIF EAGLCLIENT}


end.








