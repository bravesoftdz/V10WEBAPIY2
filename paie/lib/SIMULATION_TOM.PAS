{***********UNITE*************************************************
Auteur  ...... :  NA
Créé le ...... : 06/06/2006
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : SIMULATION (SIMULATION)
                 (modification et création interdits sur l'écran)
Mots clefs ... : TOM;SIMULATION
*****************************************************************}
Unit SIMULATION_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche, 
     fe_main,
{$else}
     eFiche, 
     maineagl,
{$ENDIF}
     ComCtrls,
     HCtrls,
     HTB97,
     Utob,
     UTOM;


Type
  TOM_SIMULATION = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    private
    procedure RESULTSIMUL(sender : TObject);
    procedure popselect(sender : TObject);
    procedure Dupliquersim(sender : TObject);
    procedure Modifiersim(sender : TObject);
    end ;

Implementation

procedure TOM_SIMULATION.OnNewRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/06/2006
Modifié le ... :   /  /    
Description .. : Suppression
Mots clefs ... : 
*****************************************************************}
procedure TOM_SIMULATION.OnDeleteRecord ;
begin
  Inherited ;
 ExecuteSql('Delete from RESULTSIMUL where PSR_CODESIMUL ="'+Getfield('PSL_CODESIMUL')+'"');
end ;


procedure TOM_SIMULATION.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_SIMULATION.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_SIMULATION.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_SIMULATION.OnLoadRecord ;
begin
  Inherited ;
  TFFiche(Ecran).Caption := 'Simulation :' + ' ' + Getfield('PSL_CODESIMUL') + ' ' + Getfield('PSL_LIBELLE') ;
 end ;

procedure TOM_SIMULATION.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/06/2006
Modifié le ... :   /  /
Description .. : On argument
Mots clefs ... :
*****************************************************************}
procedure TOM_SIMULATION.OnArgument ( S: String ) ;
var
Btn : TToolBarButton97;
begin
  Inherited ;
   Btn := TToolBarButton97(GetControl('B_DETAILSIMUL'));
  if btn <> nil then Btn.Onclick := RESULTSIMUL;
   Btn := TToolBarButton97(GetControl('B_POPSELECT'));
  if btn <> nil then Btn.Onclick := popselect;
    Btn := TToolBarButton97(GetControl('BDUPLIQUER'));
  if btn <> nil then Btn.Onclick := Dupliquersim;
    Btn := TToolBarButton97(GetControl('BMODIFIER'));
  if btn <> nil then Btn.Onclick := Modifiersim;

end ;


procedure TOM_SIMULATION.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_SIMULATION.OnCancelRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/06/2006
Modifié le ... :   /  /
Description .. : Bouton "Détail de la simulation"
Mots clefs ... :
*****************************************************************}
Procedure TOM_SIMULATION.Resultsimul(sender : TObject);
Begin
AglLancefiche('PAY', 'RESULTSIMUL_MUL', '', '',Getfield('PSL_CODESIMUL')+';'+Getfield('PSL_LIBELLE'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 01/08/2006
Modifié le ... :   /  /
Description .. : Bouton "Détail des populations sélectionnées"
Mots clefs ... :
*****************************************************************}
Procedure TOM_SIMULATION.popselect(sender : TObject);
Begin
AglLancefiche('PAY', 'POPULSIMUL_MUL', '', '',Getfield('PSL_CODESIMUL'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 18/08/2006
Modifié le ... :   /  /    
Description .. : Dupliquer une simulation
Mots clefs ... : 
*****************************************************************}
Procedure TOM_SIMULATION.Dupliquersim(sender : TObject);
var
Retour : String;
Begin
retour :=AglLancefiche('PAY', 'CALCULSIMUL', '', '','D' +';'+Getfield('PSL_CODESIMUL'));

if retour = 'OK' then TFFiche(Ecran).BFermeClick(Self);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 18/08/2006
Modifié le ... :   /  /    
Description .. : Modifier la simulation
Mots clefs ... : 
*****************************************************************}
Procedure TOM_SIMULATION.Modifiersim(sender : TObject);

Begin
AglLancefiche('PAY', 'CALCULSIMUL', '', '','M' +';'+Getfield('PSL_CODESIMUL'));
refreshdb;
end;


Initialization
  registerclasses ( [ TOM_SIMULATION ] ) ; 
end.
