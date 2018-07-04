unit UTOMEmplacement;

interface

Uses
    UTOM,Classes,sysutils, HCtrls ;

Type
    TOM_EMPLACEMENT = Class (TOM)

    Procedure OnLoadRecord ; override ;
    Procedure OnUpdateRecord ; override ;
    Procedure OnNewRecord ; override ;
    procedure OnDeleteRecord ; override ; 
    end ;

Const
	// libellés des messages
	TextMessage: array[1..5] of string 	= (
          {1}         'Vous devez renseigner un dépôt'
          {2}        ,'Vous devez renseigner un code'
          {3}        ,'Vous devez renseigner un libellé'
          {4}        ,'Vous devez renseigner un type d''emplacement'
          {5}        ,'Cet emplacement est déjà référencé dans les mouvements'
                  );


implementation

procedure TOM_EMPLACEMENT.OnDeleteRecord ;
Var CodeE,CodeD : String ;
begin
CodeE:=GetField('GEM_EMPLACEMENT') ;
CodeD:=GetField('GEM_DEPOT') ;
if CodeE='' then Exit ;
if ExisteSQL('SELECT GQ_EMPLACEMENT FROM DISPO WHERE GQ_DEPOT="'+CodeD+'" AND GQ_EMPLACEMENT="'+CodeE+'"') then
   BEGIN
   LastError:=5 ; LastErrorMsg:=TextMessage[5] ;
   END ;
  Inherited ;
end ;

Procedure TOM_EMPLACEMENT.OnNewRecord ;
begin
Inherited ;
    SetControlEnabled('GEM_DEPOT',true) ;
    SetControlEnabled('GEM_EMPLACEMENT',true) ;
end ;

Procedure TOM_EMPLACEMENT.OnLoadRecord ;
begin
Inherited ;
    if GetField('GEM_DEPOT')<>'' then
    begin
        SetControlEnabled('GEM_DEPOT',false) ;
        SetControlEnabled('GEM_EMPLACEMENT',false) ;
    end ;
end ;

Procedure TOM_EMPLACEMENT.OnUpdateRecord ;
begin
Inherited ;
SetField('GEM_EMPLACEMENT',Trim(GetControlText('GEM_EMPLACEMENT')));
    If GetField('GEM_DEPOT')='' then begin SetFocusControl('GEM_DEPOT') ; LastError:=1 ; LastErrorMsg:=TextMessage[LastError] ; exit ; end ;
    If GetField('GEM_EMPLACEMENT')='' then begin SetFocusControl('GEM_EMPLACEMENT') ; LastError:=2 ; LastErrorMsg:=TextMessage[LastError] ; exit ; end ;
    If GetField('GEM_LIBELLE')='' then begin SetFocusControl('GEM_LIBELLE') ; LastError:=3 ; LastErrorMsg:=TextMessage[LastError] ; exit ; end ;
    If GetField('GEM_TYPEEMPLACE')='' then begin SetFocusControl('GEM_TYPEEMPLACE') ; LastError:=4 ; LastErrorMsg:=TextMessage[LastError] ; exit ; end ;
end ;

Initialization
RegisterClasses([TOM_EMPLACEMENT]) ;
end.
