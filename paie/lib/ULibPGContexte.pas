{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 20/08/2007
Modifié le ... :
Description .. :
Mots clefs ... :
*****************************************************************}
unit ULibPGContexte;

interface

uses
 Classes ,
 SysUtils,
 UTOB ,  hCtrls
{$IFDEF EAGLSERVER}
 ,eSession ,
 uHTTP , uWA
{$ENDIF}
{$IFDEF CISXPGI} 
      {$IFNDEF EAGLCLIENT}
          ,Uscript
      {$ELSE}
          ,UscriptTob
      {$ENDIF}
{$ENDIF}
;

type
 TypeTPGContexte = Class of TPGContexte ;

 // Objet contentant l'ensemble des classes nécessaire au fct d'une paie eAglServeur
 TPGContexte = Class(TPersistent)
  private
    FGblTob_ParamPop,FTob_ExerciceSocial,FTob_DateClotureCp,FTob_DateCp,FTob_OrganismeBull,FTob_TauxBull   :Tob;
    FDateDebutExerSoc,FMemDateDeb,FMemDateFin,FDateJamais,FDateCp     :TDateTime;
    FMemEtab             :string;
    FCpAnnee,FCpMois,FCpJour   :Word;
    FLibBanque           :String;
    FMaxCum              :integer;
    FTabCumul : array of string;
    FTabCumulEx, FTabCumulMois: array of double;
    FTab_ZonesCalculees  :array of Double;
    FWhereCumul           :String;
    FNiveau               :String;
    FGblEditBulCp         :String;
    FAN,FPN,FRN,FBN,FPN1,FAN1,FRN1,FBN1  :double;
    FDDP,FDFP,FDDPN1,FDFPN1   :TDateTime;
    FCumul21              :String;
    Forganisme            :String;
    FCMDD, FCMDF          :TDateTime;
    FCMSal, FCMEtab       :String;
  {$IFDEF EAGLSERVER}
   FMySession   : TISession ;     // pointeur sur la session associe a l'objet ( ne sert pas pour l'instant )
{$ENDIF}
   FCount          : integer ;
   function GetTabCumul(Index: Integer): String;
   procedure SetTabCumul(Index: Integer ; Value : String);
   function GetTabCumulEx(Index: Integer): double;
   procedure SetTabCumulEx(Index: Integer ; Value : double);
   function GetTabCumulMois(Index: Integer): double;
   procedure SetTabCumulMois(Index: Integer ; Value : double);
   function GetTab_ZonesCalculees(Index: Integer): double;
   procedure SetTab_ZonesCalculees(Index: Integer ; Value : double);

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
   class function  GetCurrent  : TPGContexte ; // retourne un pointeur sur le contexte de l'utilisateur, ou le cree s'il n'existe pas
   class procedure Release ; // detruit le contexte de l'utilisateur

    property GblTob_ParamPop     :Tob        read FGblTob_ParamPop    write FGblTob_ParamPop;
    property Tob_ExerciceSocial  :Tob        read FTob_ExerciceSocial write FTob_ExerciceSocial;
    property Tob_DateClotureCp   :Tob        read FTob_DateClotureCp  write FTob_DateClotureCp;
    property Tob_DateCp          :Tob        read FTob_DateCp         write FTob_DateCp;
    property Tob_OrganismeBull   :Tob        read FTob_OrganismeBull  write FTob_OrganismeBull;
    property Tob_TauxBull        :Tob        read FTob_TauxBull       write FTob_TauxBull; 
    property DateDebutExerSoc    :TDateTime  read FDateDebutExerSoc   write FDateDebutExerSoc;
    property MemDateDeb          :TDateTime  read FMemDateDeb         write FMemDateDeb;
    property MemDateFin          :TDateTime  read FMemDateFin         write FMemDateFin;
    property DateJamais          :TDateTime  read FDateJamais         write FDateJamais;
    property DateCp              :TDateTime  read FDateCp             write FDateCp;
    property MemEtab             :string     read FMemEtab            write FMemEtab;
    property CpAnnee             :Word       read FCpAnnee            write FCpAnnee;
    property CpMois              :Word       read FCpMois             write FCpMois;
    property CpJour              :Word       read FCpJour             write FCpJour;
    property Count               :integer    read FCount              write FCount ;
    property LibBanque           :String     read FLibBanque          write FLibBanque;
    property MaxCum              :integer    read FMaxCum             write FMaxCum;
    property TabCumul[Index: Integer]: String read GetTabCumul        write SetTabCumul;default;
    property TabCumulEx[Index: Integer]: double read GetTabCumulEx    write SetTabCumulEx;
    property TabCumulMois[Index: Integer]: double read GetTabCumulMois    write SetTabCumulMois;
    property Tab_ZonesCalculees[Index: Integer]: double read GetTab_ZonesCalculees write SetTab_ZonesCalculees;
    property WhereCumul          :String     read FWhereCumul         write FWhereCumul;
    property Niveau              :String     read FNiveau             write FNiveau;
    property GblEditBulCp        :String     read FGblEditBulCp       write FGblEditBulCp;
    property AN                  :double     read FAN                 write FAN;
    property PN                  :double     read FPN                 write FPN;
    property RN                  :double     read FRN                 write FRN;
    property BN                  :double     read FBN                 write FBN;
    property PN1                 :double     read FPN1                write FPN1;
    property AN1                 :double     read FAN1                write FAN1;
    property RN1                 :double     read FRN1                write FRN1;
    property BN1                 :double     read FBN1                write FBN1;
    property DDP                 :TDateTime  read FDDP                write FDDP;
    property DFP                 :TDateTime  read FDFP                write FDFP;
    property DDPN1               :TDateTime  read FDDPN1              write FDDPN1;
    property DFPN1               :TDateTime  read FDFPN1              write FDFPN1;
    property Cumul21             :String     read FCumul21            write FCumul21;
    property organisme           :String     read Forganisme          write Forganisme;
    property CMDD                :TDateTime  read FCMDD               write FCMDD;
    property CMDF                :TDateTime  read FCMDF               write FCMDF;
    property CMSal               :String     read FCMSal              write FCMSal;
    property CMEtab              :String     read FCMEtab             write FCMEtab;
{$IFDEF EAGLSERVER}
   property MySession   : TISession     read FMySession ;
{$ENDIF}
 end ;


// fonction de recensement des classes contexte
procedure RegisterClassContext ( TheClass : TypeTPGContexte ) ; // recense un classe
function  FindClassPGContext ( const ClassName : String):  TypeTPGContexte ; // retrouve une instance de classe par son nom
function  GetClassPGContext ( TheClassName : String ) : Boolean ; // retourne true si TheClassName correspond a une classe recensée

implementation

uses
  {$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
  {$ENDIF}
  HEnt1,ParamSoc, UtilPgi;

var
 // classlist contenant les instance de classe TPGContexte enregistré
 vClassList : TStringList;

{$IFNDEF EAGLSERVER}
var
 gContexte : TPGContexte ;
{$ENDIF}



{$IFNDEF EAGLSERVER}
constructor TPGContexte.Create ;
{$ELSE}
constructor TPGContexte.Create( MySession : TISession ) ;
{$ENDIF}
begin
{$IFDEF EAGLSERVER}
 FMySession     := MySession ;
{$ENDIF}

end;

destructor TPGContexte.Destroy;
begin
 inherited ;
end;

class function TPGContexte.GetCurrent : TPGContexte ;
{$IFDEF EAGLSERVER}
var
 MySession      : TISession ;
 lIndex         : integer ;
{$ENDIF}

begin
{$IFDEF EAGLSERVER}
 // on recupere la session courante
 MySession := LookupCurrentSession ;

 // on recherche si cette classe a deja ete cree
 // on utilise le classname pour stocker/retrouver un classe
 lIndex    := MySession.UserObjects.IndexOf(ClassName) ; 

 if lIndex > -1 then
  result := TPGContexte(MySession.UserObjects.Objects[lIndex]) // on renvoie un pointeur sur la classe courante
   else
    begin
     // la classe n'existe pas il faut la creer
     // on recherche la classe ds l'ensemble des classes recensées et on la creer ds la foulée
     result       := FindClassPGContext(ClassName).Create(MySession) ;
     Result.Count := Result.Count + 1 ;
     MySession.UserObjects.AddObject(ClassName,result) ; // on ajoute la nouvelle instance de classe a la session
     if ClassParent.InheritsFrom(TPGContexte) then
      MySession.UserObjects.AddObject(ClassParent.ClassName,result) ; // on ajoute la classe ancetre a la session si elle est de type TPGContexte
    end ;

{$ELSE}
 if gContexte = nil then
  gContexte := TPGContexte.Create ;
 Result := gContexte;
{$ENDIF}

end;

class procedure TPGContexte.Release ;
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

function TPGContexte.GetTabCumul(Index: Integer): String;
begin
  Result := FTabCumul[Index];
end;

procedure TPGContexte.SetTabCumul(Index: Integer ; Value : String);
begin
  if Length(FTabCumul) < (Index + 1) then
    SetLength(FTabCumul, Index + 1);
  FTabCumul[Index] := Value;
end;

function TPGContexte.GetTabCumulEx(Index: Integer): double;
begin
  Result := FTabCumulEx[Index];
end;

procedure TPGContexte.SetTabCumulEx(Index: Integer ; Value : double);
begin
  if Length(FTabCumulEx) < (Index + 1) then
    SetLength(FTabCumulEx, Index + 1);
  FTabCumulEx[Index] := Value;
end;

function TPGContexte.GetTabCumulMois(Index: Integer): double;
begin
  Result := FTabCumulMois[Index];
end;

procedure TPGContexte.SetTabCumulMois(Index: Integer ; Value : double);
begin
  if Length(FTabCumulMois) < (Index + 1) then
    SetLength(FTabCumulMois, Index + 1);
  FTabCumulMois[Index] := Value;
end;

function TPGContexte.GetTab_ZonesCalculees(Index: Integer): double;
begin
  Result := FTab_ZonesCalculees[Index];
end;

procedure TPGContexte.SetTab_ZonesCalculees(Index: Integer ; Value : double);
begin
  if Length(FTab_ZonesCalculees) < (Index + 1) then
    SetLength(FTab_ZonesCalculees, Index + 1);
  FTab_ZonesCalculees[Index] := Value;
end;

function  GetClassPGContext ( TheClassName : String ) : Boolean ;
begin
 Result := vClassList.IndexOf( UpperCase( TheClassName ) ) <> -1;
end ;

procedure RegisterClassContext ( TheClass : TypeTPGContexte );
begin
 if not GetClassPGContext(TheClass.ClassName) then vClassList.AddObject( TheClass.ClassName, TObject(TheClass) );
end ;

function  FindClassPGContext ( const ClassName : String):  TypeTPGContexte ;
var
  i : Integer;
begin
  i := vClassList.IndexOf(ClassName) ;
  if i = -1 then
   result := nil
    else
     Result := TypeTPGContexte(vClassList.Objects[i]) ;
end;



initialization
 vClassList := TStringList.Create;
 RegisterClassContext(TPGContexte) ;


finalization
 vClassList.Free;

end.

