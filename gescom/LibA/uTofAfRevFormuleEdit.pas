{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 08/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFREVFORMULEEDIT ()
Mots clefs ... : TOF;AFREVFORMULEEDIT
*****************************************************************}
Unit uTofAfRevFormuleEdit;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$Else}
     MainEagl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1, UTof, UtilRevFormule;

Type
  TOF_AFREVFORMULEEDIT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    private
      fStFormule : String;
      fStAffaire : String;
      fStForCode : String;
  end;

procedure AglLanceFicheAFREVFORMULEEDIT(cle,Action : string ) ;

Implementation

procedure TOF_AFREVFORMULEEDIT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFREVFORMULEEDIT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFREVFORMULEEDIT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFREVFORMULEEDIT.OnLoad ;
var
  vMemoEdition  : TMemo;
  vMemoDetail   : TMemo;
  
begin

  vMemoDetail := TMemo.create(Ecran);
  vMemoDetail.Parent := Ecran;
  vMemoEdition := TMemo(GetControl('MEMO'));

  try
    if not rien(fStFormule) then
    begin                                     
      RemplaceIndice(fStForCode, fStAffaire, fStFormule);
      TraitementFormule(fStFormule, vMemoDetail, vMemoEdition);
    end;
  finally
    vMemoDetail.Free;
  end;
end;
 
procedure TOF_AFREVFORMULEEDIT.OnArgument (S : String ) ;
Var
  Critere, Champ, valeur  : String;
  X : integer ;

begin
  Inherited ;
  fStAffaire := '';
  Critere:=(Trim(ReadTokenSt(S)));
  While (Critere <>'') do
  Begin
    if Critere<>'' then
    Begin
      X:=pos('=',Critere);
        if x<>0 then
        begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
        end;
        if (Champ ='FORMULE') then fStFormule := valeur
        else if (Champ ='AFFAIRE') then fStAffaire := valeur
        else if (Champ ='FORCODE') then fStForCode := valeur;
    END;                         
    Critere:=(Trim(ReadTokenSt(S)));
  END;
end;

procedure TOF_AFREVFORMULEEDIT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AFREVFORMULEEDIT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFREVFORMULEEDIT.OnCancel () ;
begin
  Inherited ;
end ;

procedure AglLanceFicheAFREVFORMULEEDIT(cle,Action : string ) ;
begin
  AglLanceFiche('AFF','AFREVFORMULEEDIT','',cle,Action);
end ;
                                          
Initialization
  registerclasses ( [ TOF_AFREVFORMULEEDIT ] ) ;
end.
