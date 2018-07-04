{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 20/12/2000
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : TACHE (TACHE)
Mots clefs ... : TOM;TACHE
*****************************************************************}
Unit UTOMTACHE ;

Interface

Uses StdCtrls, Controls, Classes,  forms, sysutils, ComCtrls,
{$IFDEF EAGLCLIENT}
   efiche,
{$ELSE}
   dbTables, db, Fiche,
{$ENDIF}
     CalcOleGenericAff,
     HCtrls, HEnt1, HMsgBox, UTOM,  FichList, UTob,affaireUtil;

Type T_TypeSaisieTache = (TttAffaire,TttPiece,TttLigne);

Type
  TOM_TACHE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( Stargument: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;

    private
    Affaire,Tiers,Aff0,Aff1,Aff2,Aff3,Avenant : String;
    SaisieTache :  T_TypeSaisieTache;
    end ;

Implementation

procedure TOM_TACHE.OnNewRecord ;
Var QQ : TQuery ;
    IMax :integer ;
begin
Inherited;
QQ:=OpenSQL('SELECT MAX(ATA_NUMEROTACHE) FROM TACHE',TRUE) ;
if Not QQ.EOF then imax:=QQ.Fields[0].AsInteger+1 else iMax:=1;
Ferme(QQ) ;
SetField('ATA_NUMEROTACHE',IMax);
SetField ('ATA_DATEDEBUT',idate1900); SetField ('ATA_DATEFIN',idate2099);
SetField('ATA_TIERS',Tiers);
if (SaisieTache = TttAffaire) and (Affaire <> '') then
   BEGIN
   SetField('ATA_AFFAIRE',Affaire); SetField('ATA_AFFAIRE1',Aff1);
   SetField('ATA_AFFAIRE2',Aff2); SetField('ATA_AFFAIRE3',Aff3);
   SetField('ATA_AVENANT',Avenant);
   END;
end;

procedure TOM_TACHE.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_TACHE.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_TACHE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_TACHE.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_TACHE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_TACHE.OnArgument ( Stargument: String ) ;
Var X:integer;
    Tmp,champ, valeur : string;
begin
  Inherited ;
Affaire := ''; Tiers := '';
SaisieTache := TttAffaire ;

// traitement des arguments
Tmp:=(Trim(ReadTokenSt(stArgument)));
While (Tmp <>'') do
    BEGIN
    if Tmp<>'' then
        BEGIN
        X:=pos(':',Tmp);
        if x = 0 then X:=pos('=',Tmp);
        if x<>0 then
           begin
           Champ:=copy(Tmp,1,X-1);
           Valeur:=Copy (Tmp,X+1,length(Tmp)-X);
           end
        else Champ := Tmp;
        if Champ='BYAFFAIRE'  then SaisieTache := TttAffaire else
        if Champ= 'ATA_AFFAIRE'  then Affaire:=valeur else
        if Champ= 'ATA_TIERS'    then Tiers := valeur else
        END;
    Tmp:=(Trim(ReadTokenSt(stArgument)));
    END;
// Paramétrage Affaire + Tiers
if (SaisieTache = TttAffaire) and (Affaire <> '') then
   BEGIN
   CodeAffaireDecoupe(Affaire,Aff0,Aff1,Aff2,Aff3,Avenant,taModif,false);
   ChargeCleAffaire(nil, THEdit(GetControl('CATA_AFFAIRE1')),THEdit(GetControl('CATA_AFFAIRE2')),THEdit(GetControl('CATA_AFFAIRE3')),THEdit(GetControl('CATA_AVENANT')),
                    Nil,taConsult,Affaire,False);
   END;
END;

procedure TOM_TACHE.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_TACHE.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_TACHE ] ) ; 
end.

