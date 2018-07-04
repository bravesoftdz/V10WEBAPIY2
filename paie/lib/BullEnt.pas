unit BullEnt;

interface

uses Dialogs, Classes, StdCtrls, ComCtrls, Forms, SysUtils,
     UTob, Hent1, HCtrls, Paramsoc, Controls, hmsgbox
{$IFDEF EAGLCLIENT}

{$ELSE}
  {$IFNDEF DBXPRESS} ,dbtables {$ELSE} ,uDbxDataSet {$ENDIF}
{$ENDIF}

{$IFDEF EAGLSERVER}
  ,uLibPGContexte
  ,eSession
{$ENDIF}

{$IFDEF VER150}
  ,Variants
{$ENDIF}

     {$IFDEF SERIE1}
      {$IFDEF eAGLServer}
      {$ELSE}
      ,UtTX_
      {$ENDIF !eAGLServer}
      ,S1Util, Ut
     {$ELSE}
     {$ENDIF}
     {$IFDEF A_SUPPRIMER?}
     , Forms
     , Classes
     {$ENDIF !A_SUPPRIMER?}
     ;


{$IFDEF EAGLSERVER}
{$IFDEF SERIE1}
type
   TPGContexte = class(TObject)
  protected { Déclarations protégées}
     constructor Create( MySession : TISession ) ; virtual;
  public
     destructor  Destroy ; override ;
     class function  GetCurrent  : TPGContexte ; // retourne un pointeur sur le contexte de l'utilisateur, ou le cree s'il n'existe pas
     class procedure Release ; // detruit le contexte de l'utilisateur

   end ;
{$ENDIF}

type
  TEditionPaieContext = class (TPGContexte)

  private { Déclarations privées}
    SavCalcEdt: TProcCalcEdt;

  protected { Déclarations protégées}
    constructor Create( MySession : TISession ) ; override;

  public { Déclarations publiques}
    destructor Destroy; override;
  end;
{$ENDIF}

implementation


{ TEditionPaieContext }

{$IFDEF EAGLSERVER}
{$IFDEF SERIE1}
constructor TPGContexte.Create( MySession : TISession ) ;
begin
end;

destructor TPGContexte.Destroy;
begin
 inherited ;
end;

class function TPGContexte.GetCurrent : TPGContexte ;
var
 MySession      : TISession ;
 lIndex         : integer ;
begin

 // on recupere la session courante
 MySession := LookupCurrentSession ;

 // on recherche si cette classe a deja ete cree
 // on utilise le classname pour stocker/retrouver un classe
 lIndex    := MySession.UserObjects.IndexOf('ID_PAIE') ;

 if lIndex > -1 then
  result := TPGContexte(MySession.UserObjects.Objects[lIndex]) // on renvoie un pointeur sur la classe courante
   else
    begin
     // la classe n'existe pas il faut la creer
     // on recherche la classe ds l'ensemble des classes recensées et on la creer ds la foulée
     result       := TEditionPaieContext.Create(MySession) ;

     TEditionPaieContext(result).SavCalcEdt:=LookUpCurrentSession.UserCalcEdt;

     MySession.UserObjects.AddObject('ID_PAIE',result) ; // on ajoute la nouvelle instance de classe a la session
     if ClassParent.InheritsFrom(TPGContexte) then
      MySession.UserObjects.AddObject(ClassParent.ClassName,result) ; // on ajoute la classe ancetre a la session si elle est de type TPGContexte
    end ;


end;

class procedure TPGContexte.Release ;
var
 MySession : TISession ;
 lIndex    : integer ;
 MyContext : tObject ;
begin
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
end ;
{$ENDIF}

constructor TEditionPaieContext.Create ( MySession : TISession );
begin
  inherited  Create(MySession);
end;

destructor TEditionPaieContext.Destroy;
begin
  LookUpCurrentSession.UserCalcEdt:=SavCalcEdt;

  inherited;
end;
{$ENDIF}


initialization

finalization

end.






