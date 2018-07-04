{***********UNITE*************************************************
Auteur  ...... : Manou Entressangle & Laurent Gendreau
Créé le ...... : 15/04/2005
Modifié le ... : 18/04/2005
Description .. :
Mots clefs ... :
*****************************************************************}
unit ULibCpContexte;

interface

uses
 Classes ,
 SysUtils,
 ULibExercice ,
 UTOB ,  hCtrls,
{$IFDEF EAGLSERVER}
 eSession ,
 uHTTP , uWA ,
{$ENDIF}
{$IFDEF CISXPGI} // A remplacer plus tard par CISXPGI
      {$IFNDEF EAGLCLIENT}
          Uscript,
      {$ELSE}
          UscriptTob,
      {$ENDIF}
{$ENDIF}
{$IFDEF MODENT1}
  CPObjetGen,
{$ENDIF MODENT1}
 Ent1 ;


type


 TypeTCPContexte = Class of TCPContexte ;

 // Objet contentant l'ensemble des classes nécessaire au fct d'une compta eAglServeur
 TCPContexte = Class(TPersistent)
  private
   FZExercice      : TZexercice ;     // Objet de gestion des exercices ( ULibEcriture.pas )
   FZInfoCpta      : TZInfoCpta ;    // Objet de gestion des Info Compta ( longueur de compte etc ds Ent1 )
   FZHalleyUser    : TZHalleyUser ;  // Objet des gestion des utilisateurs ( ds Ent1 )
   FZCatBud        : TZCatBud;       {Objet contenant les catégories budgétaires}
{$IFDEF EAGLSERVER}
   FMySession   : TISession ;     // pointeur sur la session associe a l'objet ( ne sert pas pour l'instant )
{$ENDIF}
   FLaTOBTVA       : TOB ;
   FCount          : integer ;
   FTresoNoDossier : string;
{$IFDEF CISXPGI}
   FVarCisx        : TZVarCisx ;    // les variables Cisx
{$ENDIF}

   function GetTresoNoDossier : string;
  protected
   // les methode de création/destruction de l'objet ne sont pas visible de l'exterieur
   // on ne doit jamais les appeler
{$IFNDEF EAGLSERVER}
       constructor Create ; virtual ;
{$ELSE}
       constructor Create( MySession : TISession ) ; virtual ;
{$ENDIF}

   destructor  Destroy ; override ;

  public
   // fonction d'acces de création et de destruction de l'objet
   class function  GetCurrent  : TCPContexte ; // retourne un pointeur sur le contexte de l'utilisateur, ou le cree s'il n'existe pas
   class procedure Release ; // detruit le contexte de l'utilisateur

   property Exercice       : TZexercice    read FZExercice    write FZExercice ;
   property InfoCpta       : TZInfoCpta    read FZInfoCpta    write FZInfoCpta ;
   property HalleyUser     : TZHalleyUser  read FZHalleyUser  write FZHalleyUser ;
   property InfoCatBud     : TZCatBud      read FZCatBud      write FZCatBud;
   property LaTOBTVA       : TOB           read FLaTOBTVA     write FLaTOBTVA ;
   property Count          : integer       read FCount        write FCount ;
{$IFDEF EAGLSERVER}   property MySession   : TISession     read FMySession ; {$ENDIF}
   property TresoNoDossier : string        read GetTresoNoDossier;
{$IFDEF CISXPGI}
   property VarCisx        : TZVarCisx    read FVarCisx    write FVarCisx ;
{$ENDIF}
 end ;


// fonction de recensement des classes contexte
procedure RegisterClassContext ( TheClass : TypeTCPContexte ) ; // recense un classe
function  FindClassCPContext ( const ClassName : String):  TypeTCPContexte ; // retrouve une instance de classe par son nom
function  GetClassCPContext ( TheClassName : String ) : Boolean ; // retourne true si TheClassName correspond a une classe recensée

implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  {$ENDIF MODENT1}
  {$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
  {$ENDIF}
  HEnt1,ParamSoc, UtilPgi;

var
 // classlist contenant les instance de classe TCpContexte enregistré
 vClassList : TStringList;

{$IFNDEF EAGLSERVER}
var
 gContexte : TCpContexte ;
{$ENDIF}



{$IFNDEF EAGLSERVER}
constructor TCPContexte.Create ;
{$ELSE}
constructor TCPContexte.Create( MySession : TISession ) ;
{$ENDIF}
begin
{$IFDEF EAGLSERVER}
 FMySession     := MySession ;
{$ENDIF}

 FZExercice     := TZExercice.Create ;
 FZInfoCpta     := TZInfoCpta.create;
 FZHalleyUser   := TZHalleyUser.Create ;
{$IFDEF CISXPGI}
 VarCisx        := TZVarCisx.create;
{$ENDIF}
 FLaTOBTVA      := TOB.Create('', nil , -1) ;
 FLaTOBTVA.LoadDetailDB('TXCPTTVA','','',nil,True) ;
// FZc
end;

destructor TCPContexte.Destroy;
begin
 FreeAndNil(FZExercice);
 FreeAndNil(FZInfoCpta);
 FreeAndNil(FZHalleyUser);
 FreeAndNil(FLaTOBTVA);
{$IFDEF CISXPGI}
 FreeAndNil (FVarCisx);
{$ENDIF}
 inherited ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : JP
Créé le ...... : 02/10/06
Modifié le ... : 02/10/06
Description .. : Chargement du NoDossier de la base de Tréso
Mots clefs ... :
*****************************************************************}
function TCPContexte.GetTresoNoDossier : string;
var
  Q : TQuery;
  T : Boolean;
begin
  if FTresoNoDossier = '' then begin
    T := (GetParamSocSecur('SO_TRBASETRESO', '') <> V_PGI.SchemaName) and (GetParamSocSecur('SO_TRBASETRESO', '') <> '');
    //Si la Tréso est multi sociétés et que l'on n'est pas sur la base de Treso ...
    if EstComptaTreso and EstMultiSoc and T then begin
      //... On recherche le NoDossier de la base Tréso
      Q := OpenSQL('SELECT DOS_NODOSSIER FROM DOSSIER WHERE DOS_NOMBASE = "' + GetParamSocSecur('SO_TRBASETRESO', '') + '"', True);
      if not Q.EOF then
        FTresoNoDossier := Q.FindField('DOS_NODOSSIER').AsString;
      Ferme(Q);
    end
    else begin
      ChargeNoDossier;
      FTresoNoDossier := V_PGI.NoDossier;
    end;
  end;
  Result := FTresoNoDossier;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Manou Entressangle & Laurent Gendreau
Créé le ...... : 21/04/2005
Modifié le ... : 21/04/2005
Description .. : Methode d'acces unique a une classe TCpContexte.
Mots clefs ... : 
*****************************************************************}
class function TCPContexte.GetCurrent : TCPContexte ;
{$IFDEF EAGLSERVER}
var
 MySession      : TISession ;
 lIndex         : integer ;
{$ENDIF}

begin
{$IFDEF TTW}
 cWA.MessagesAuClient('COMSX.IMPORT','','TCPContexte.GetCurrent') ;
{$ENDIF}

{$IFDEF EAGLSERVER}
 // on recupere la session courante
 MySession := LookupCurrentSession ;

 // on recherche si cette classe a deja ete cree
 // on utilise le classname pour stocker/retrouver un classe
 lIndex    := MySession.UserObjects.IndexOf(ClassName) ; 

 if lIndex > -1 then
  result := TCPContexte(MySession.UserObjects.Objects[lIndex]) // on renvoie un pointeur sur la classe courante
   else
    begin
     {$IFDEF TTW}
      cWA.MessagesAuClient('COMSX.IMPORT','','TCPContexte.GetCurrent : creation de la classe ' + ClassName) ;
     {$ENDIF}   
     // la classe n'existe pas il faut la creer
     // on recherche la classe ds l'ensemble des classes recensées et on la creer ds la foulée
     result       := FindClassCPContext(ClassName).Create(MySession) ;
     Result.Count := Result.Count + 1 ;
     MySession.UserObjects.AddObject(ClassName,result) ; // on ajoute la nouvelle instance de classe a la session
     if ClassParent.InheritsFrom(TCPContexte) then
      MySession.UserObjects.AddObject(ClassParent.ClassName,result) ; // on ajoute la classe ancetre a la session si elle est de type TCpContexte
    end ;

{$ELSE}
 if gContexte = nil then
  gContexte := TCPContexte.Create ;
 Result := gContexte;
{$ENDIF}

end;


{***********A.G.L.***********************************************
Auteur  ...... : Manou Entressangle & Laurent Gendreau
Créé le ...... : 21/04/2005
Modifié le ... : 21/04/2005
Description .. : Methode de destruction du contexte courant
Mots clefs ... :
*****************************************************************}
class procedure TCPContexte.Release ;
{$IFDEF EAGLSERVER}
var
 MySession : TISession ;
 lIndex    : integer ;
 MyContext : tObject ;
{$ENDIF}
begin
{$IFDEF EAGLSERVER}
 // on recupere le pointeur sur la session courante
 MySession := LookupCurrentSession ;
 // on recherche la classe ds la liste des objets de l'utilisateurs
 lIndex    := MySession.UserObjects.IndexOf(ClassName) ;
 if lIndex > -1 then
  begin
   // on detruit le contexte utilisateur
   MyContext := MySession.UserObjects.Objects[lIndex] ;
   MySession.UserObjects.Delete(lIndex) ;
   FreeAndNil(MyContext) ;
  end ;
 // on recherche si on avait aussi ajouter la classe ancetre a la liste des objets utilisateur
 lIndex := MySession.UserObjects.IndexOf(ClassParent.ClassName) ;
 if lIndex > -1 then
  MySession.UserObjects.Delete(lIndex) ;
{$ELSE}
 gContexte.free ;
 gContexte := nil ;
{$ENDIF}
end ;

function  GetClassCPContext ( TheClassName : String ) : Boolean ;
begin
 Result := vClassList.IndexOf( UpperCase( TheClassName ) ) <> -1;
end ;

procedure RegisterClassContext ( TheClass : TypeTCPContexte );
begin
 if not GetClassCPContext(TheClass.ClassName) then vClassList.AddObject( TheClass.ClassName, TObject(TheClass) );
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Manou Entressangle & Laurent Gendreau
Créé le ...... : 21/04/2005
Modifié le ... : 21/04/2005
Description .. : retrouve une classe recense par son nom
Mots clefs ... : 
*****************************************************************}
function  FindClassCPContext ( const ClassName : String):  TypeTCPContexte ;
var
  i : Integer;
begin
  i := vClassList.IndexOf(ClassName) ;
  if i = -1 then
   result := nil
    else
     Result := TypeTCPContexte(vClassList.Objects[i]) ;
end;


initialization
 vClassList := TStringList.Create;
 RegisterClassContext(TCPContexte) ;


finalization
 vClassList.Free;

end.
