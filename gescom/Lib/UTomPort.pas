unit UTomPort;

interface
Uses
    UTOM,Classes,sysutils,Forms,
{$IFDEF EAGLCLIENT}
    HDB,
{$ELSE}
    DBctrls,
{$ENDIF}
    Graphics,Hent1,HCtrls ;

Type
    TOM_PORT = Class (TOM)

    Procedure OnLoadRecord ; override ;
    Procedure OnUpdateRecord ; override ;
    Procedure OnNewRecord ; override ;
    Procedure OnDeleteRecord ; override ;
end ;

Const
	// libellés des messages
	TextMessage: array[1..4] of string 	= (
          {1}         'Le code port est obligatoire.'
          {2}        ,'Le libellé est obligatoire.'
          {3}        ,'Le type de port est obligatoire.'
          {4}        ,'Vous ne pouvez pas effacer ce port, il est actuellement utilisé.'

                  );



implementation

Procedure TOM_PORT.OnNewRecord ;
begin
Inherited ;
SetField('GPO_CREATEUR',V_PGI.User);
SetField('GPO_CREERPAR','SAI');
SetField('GPO_DATESUP', iDate2099);
end;

Procedure TOM_PORT.OnLoadRecord ;
begin
Inherited ;
SetControlEnabled('GPO_PVHT',True);
SetControlEnabled('GPO_COEFF',True);
SetControlEnabled('GPO_VERROU',True); 
SetControlEnabled('GPO_PVTTC',True);
SetControlEnabled('GPO_MINIMUM',True);
SetControlProperty('GPO_PVHT','Color',clWindow);
SetControlProperty('GPO_PVTTC','Color',clWindow);
SetControlProperty('GPO_COEFF','Color',clWindow);
SetControlProperty('GPO_MINIMUM','Color',clWindow);
If GetField('GPO_TYPEPORT')='MT' then
    begin
    SetControlEnabled('GPO_COEFF',False);
    SetControlEnabled('GPO_VERROU',False); 
    SetControlProperty('GPO_COEFF','Color',clbtnFace);
    end;
If GetField('GPO_TYPEPORT')='HT' then
    begin
    SetControlEnabled('GPO_PVHT',False);
    SetControlEnabled('GPO_PVTTC',false);
    SetControlProperty('GPO_PVHT','Color',clbtnFace);
    SetControlProperty('GPO_PVTTC','Color',clbtnFace);
    end;

if TDBCheckBox(ecran.FindComponent('GPO_FRANCO')).Checked=False then
   begin
   SetControlEnabled('GPO_MINIMUM',False);
   SetControlProperty('GPO_MINIMUM','Color',clbtnFace);
   end ;
end ;

Procedure TOM_PORT.OnUpdateRecord ;
begin
Inherited ;
SetField('GPO_CODEPORT',Trim(GetControlText('GPO_CODEPORT')));
If GetField('GPO_CODEPORT')='' then
   begin
   SetFocusControl('GPO_CODEPORT') ;
   LastError:=1 ;
   LastErrorMsg:=TextMessage[LastError] ;
   exit ;
   end ;
If GetField('GPO_LIBELLE')='' then
   begin
   SetFocusControl('GPO_LIBELLE') ;
   LastError:=2 ;
   LastErrorMsg:=TextMessage[LastError] ;
   exit ;
   end ;
If GetField('GPO_TYPEPORT')='' then
   begin
   SetFocusControl('GPO_TYPEPORT') ;
   LastError:=3 ;
   LastErrorMsg:=TextMessage[LastError] ;
   exit ;
   end ;
If GetField('GPO_TYPEPORT')='MT' then
   begin
   Setfield('GPO_COEFF',0) ;
   end ;
If GetField('GPO_TYPEPORT')='HT' then
   begin
   Setfield('GPO_PVHT',0) ;
   Setfield('GPO_PVTTC',0) ;
   end ;

if TDBCheckBox(ecran.FindComponent('GPO_FRANCO')).Checked=False then
   begin
   Setfield('GPO_MINIMUM',0) ;
   end ;
end ;

procedure TOM_Port.OnDeleteRecord  ;
begin
Inherited ;

if ExisteSQL('SELECT GPT_CODEPORT FROM PIEDPORT WHERE GPT_CODEPORT="'
     +GetField('GPO_CODEPORT')+'"') then
   BEGIN
   LastError:=4;
   LastErrorMsg:=TextMessage[LastError];
   exit ;
   end ;
end;

Initialization
RegisterClasses([TOM_PORT]) ;

end.