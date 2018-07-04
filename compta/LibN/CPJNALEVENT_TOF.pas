{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 16/10/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPJNALEVENT ()
Mots clefs ... : TOF;CPJNALEVENT
*****************************************************************}
Unit CPJNALEVENT_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     HDB,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     FE_Main,     // AglLanceFiche
  {$ELSE}
     MainEagl,                        // AglLanceFiche
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,   
     HTB97,
     UTOF ;

function CPLanceFicheCPJNALEVENT( vStRange, vStLequel, vStArgs : string ) : string;

Type
  TOF_CPJNALEVENT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
{$IFDEF EAGLCLIENT}
    FListe : THGrid;
{$ELSE}
    FListe : THDBGrid;
{$ENDIF}
{$IFDEF EAGLCLIENT}
    procedure IsLineCloture ( Sender : TObject; ou : Longint; var Cancel : Boolean; Chg : Boolean);
{$ELSE}
    procedure IsLineCloture ( Sender : TObject);
{$ENDIF}                                    
    procedure VerifChecksum ( Sender : TObject );   
    function  IsTypeCloture : Boolean;
    function  CalculCheckSum ( Session : integer ) : String;
    function  GetCS( BlocNote : HTStringList ) : String;
  end ;

Implementation

uses
  Grids
  ,TntDBGrids
  ,DBGrids
  ,Variants
  ,TntWideStrings
  ,uLibValidation;

function CPLanceFicheCPJNALEVENT( vStRange, vStLequel, vStArgs : string ) : string;
begin
  result := AGLLanceFiche('CP', 'CPJNALEVENT', vStRange, vStLequel, vStArgs ) ;
end ;

procedure TOF_CPJNALEVENT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPJNALEVENT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPJNALEVENT.OnUpdate ;   
{$IFDEF EAGLCLIENT}
var
  cancel : boolean;
{$ENDIF}
begin
  Inherited ;       
{$IFDEF EAGLCLIENT}
  IsLineCloture(nil,-1,cancel,false);
{$ELSE}
  IsLineCloture(nil);
{$ENDIF}
end ;

procedure TOF_CPJNALEVENT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPJNALEVENT.OnArgument (S : String ) ;
begin
  Inherited ;

  // Initialisation composants
{$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FListe'));
{$ELSE}
  FListe := THDBGrid(GetControl('Fliste'));
{$ENDIF}

  // Initialisation evenements
  TToolbarButton97(GetControl('BCHECK')).OnClick := VerifChecksum;     
  FListe.OnDblClick := VerifChecksum;
  FListe.OnRowExit := IsLineCloture;
end ;

procedure TOF_CPJNALEVENT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPJNALEVENT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPJNALEVENT.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/10/2007
Modifié le ... :   /  /    
Description .. : Permet d'afficher un message indiquant si le checksum 
Suite ........ : present est correct avec celui recalculé ou non
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPJNALEVENT.VerifChecksum ( Sender : TObject );
var
  Event : integer;
  BlocNote : HTStringList;
  CChecksum : String;
  LChecksum : String;
begin
  if IsTypeCloture then
  begin
     // On est sur un evenement de type cloture
     BlocNote := HTStringList.Create;
     try
        BlocNote.Add(GetField('GEV_BLOCNOTE'));
        Event := GetField('GEV_NUMEVENT');
        CChecksum := CalculCheckSum(Event);
        LChecksum := GetCS(BlocNote);
        if CChecksum <> LChecksum then
           PGIError('Le checksum calculé est différent de celui present dans la trace')
        else
           PGIInfo('Les checksums sont identiques');
     finally
        BlocNote.Free;
     end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/10/2007
Modifié le ... :   /  /    
Description .. : Permet d'afficher ou non le bouton permettant le calcul du 
Suite ........ : checksum
Mots clefs ... : 
*****************************************************************}
{$IFDEF EAGLCLIENT}
procedure TOF_CPJNALEVENT.IsLineCloture (Sender : TObject; ou : Longint; var Cancel : Boolean; Chg : Boolean);
{$ELSE}
procedure TOF_CPJNALEVENT.IsLineCloture (Sender : TObject);
{$ENDIF}
begin
   SetControlVisible('BCHECK',IsTypeCloture);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/10/2007
Modifié le ... :   /  /
Description .. : Indique si la ligne courante est de type cloture 
Mots clefs ... : 
*****************************************************************}
function TOF_CPJNALEVENT.IsTypeCloture : Boolean;
begin          
  Result := false;
  if VarIsNull(GetField('GEV_TYPEEVENT')) then Exit ;
  if ( GetField('GEV_TYPEEVENT') = 'CPE' ) or ( GetField('GEV_TYPEEVENT') = 'CEX' ) then
          Result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/10/2007
Modifié le ... :   /  /    
Description .. : Retourne le checksum correspondant aux ecritures validées 
Suite ........ : lors de la session passée en paramètre
Mots clefs ... : 
*****************************************************************}
function  TOF_CPJNALEVENT.CalculCheckSum ( Session : integer ) : String;
var
  SQL : String;
  Q   : TQuery;
begin
  Result := '';
  // On recupere d'abord les id de validation
  SQL := 'SELECT CPV_IDDEBUTVAL,CPV_IDFINVAL FROM CPJALVALIDATION WHERE CPV_NUMEVENT = "' + IntToStr(Session) + '"';
  Q := OpenSQL(SQL,true);
  try
     if not Q.Eof then
     begin
        SQL := 'E_DOCID >= "' + Q.FindField('CPV_IDDEBUTVAL').AsString + '" AND E_DOCID <= "' + Q.FindField('CPV_IDFINVAL').AsString + '"';
        Result := GetChecksumEcriture(SQL + ' AND (E_GENERAL LIKE "1%" OR E_GENERAL LIKE "2%" OR E_GENERAL LIKE "3%" OR E_GENERAL LIKE "4%" OR E_GENERAL LIKE "5%")');
        Result := Result + GetChecksumEcriture(SQL + ' AND (E_GENERAL LIKE "6%" OR E_GENERAL LIKE "7%")');
     end;
  finally
     Ferme(Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/10/2007
Modifié le ... :   /  /    
Description .. : Récupere le checksum du blocnote
Mots clefs ... : 
*****************************************************************}
function  TOF_CPJNALEVENT.GetCS( BlocNote : HTStringList ) : String;
var
  i,j : integer;
  temp : string;
begin
  Result := '';
  // On commence par la fin car c'est generalement la derniere chose inscrite
  for i := BlocNote.Count - 1 downto 0 do
  begin
     j := System.Pos('CHECKSUM=',BlocNote[i]);
     if (j > 0) then
     begin
        Temp := System.Copy(BlocNote[i],j + 9,Length(BlocNote[i]) - j - 8);
        j := System.Pos(' ',Temp);
        if (j > 0) then
           Result := System.Copy(Temp,1,j - 1)
        else
           Result := Temp;
        Break;
     end;
  end;
end;

Initialization
  registerclasses ( [ TOF_CPJNALEVENT ] ) ; 
end.

