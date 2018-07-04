unit UTomFerie;

interface
uses sysutils,Hent1,Classes,UTOM,
{$IFDEF EAGLCLIENT}
     Spin,Maineagl,
{$ELSE}
     db,HDB,FE_Main,
{$ENDIF}
      UTOB, StdCtrls;

Type
     TOM_JourFerie = Class (TOM)
      procedure OnUpdateRecord; override;
       procedure OnLoadRecord; override;
       procedure OnNewRecord  ; override ;
       procedure OnArgument (stArgument : String ) ; override ;
       Procedure ChangeLibMois (Sender: TObject);
       private
{$IFDEF EAGLCLIENT}
      CCMois: TSpinEdit;
{$ELSE}
      CCMois: THDBSpinEdit;
{$ENDIF} 
     End;
 Procedure AFLanceFiche_JourFerie;

implementation

{TOM Jour Fériés}
Procedure TOM_JourFerie.OnArgument (stArgument : String ) ;
BEGIN
Inherited;
{$IFDEF EAGLCLIENT}
CCMois := TSpinEdit(GetControl('AJF_MOIS'));
{$ELSE}
CCMois := THDBSpinEdit(GetControl('AJF_MOIS'));
{$ENDIF}
if CCMois <> Nil then CCMois.OnChange := ChangeLibMois; 

END;

Procedure TOM_JourFerie.OnLoadRecord;
BEGIN
Inherited;
if CCMois <> nil then
  { SB 07/02/2003 Test si LibMois <> CCMois.Value }
  if GetControlText('LIBMOIS')<>LongMonthNames[CCMois.Value] then
    SetControlText ('LIBMOIS',LongMonthNames[CCMois.Value]);  
END;

Procedure TOM_JourFerie.ChangeLibMois (Sender: TObject);
BEGIN
Inherited;
if CCMois <> Nil then
  { SB 07/02/2003 Test si LibMois <> CCMois.Value }
  if GetControlText('LIBMOIS')<>LongMonthNames[CCMois.Value] then
     SetControlText ('LIBMOIS',LongMonthNames[CCMois.Value]); 
END;





Procedure TOM_JourFerie.OnNewRecord;
Var Day,Month,Year:Word;
Begin
Inherited;
SetField('AJF_JOUR',1);
SetField('AJF_MOIS',1);
DecodeDate(Now,Year,Month,Day);
SetField('AJF_ANNEE',Integer(Year));
SetField ('AJF_PREDEFINI','DOS');
SetField ('AJF_NODOSSIER',V_PGI.Nodossier);
End;

Procedure TOM_JourFerie.OnUpdateRecord;
Begin
Inherited;
SetField ('AJF_CODEFERIE', Trim (GetField('AJF_CODEFERIE')) );
If (Not IsValidDate(IntToStr(GetField('AJF_JOUR'))+'/'+IntToStr(GetField('AJF_MOIS'))+'/'+IntToStr(GetField('AJF_ANNEE')))) then
   Begin
   LastError:=3 ; LastErrorMsg:='Date incorrecte veuillez la modifier';
   End;
End;
 Procedure AFLanceFiche_JourFerie;
begin
AGLLanceFiche ('AFF','JOURFERIE','','','');
end;


Initialization
registerclasses([TOM_JourFerie]);
end.
 