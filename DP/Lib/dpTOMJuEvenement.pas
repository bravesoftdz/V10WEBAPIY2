unit dpTOMJuEvenement;
// TOM de la table JUEVENEMENT

interface

uses
     sysutils, Classes, hctrls, controls,
     UTOMEvenement;

Type
  TOM_JUEvenement = Class (TOM_Evenement)
    procedure OnArgument(stArgument: String); override;
    procedure OnNewRecord; override;
    procedure DEST_OnExit(Sender: TObject);
  private
  end;

///////////// IMPLEMENTATION ////////////

implementation

{ TOM_JUEvenement }

uses AnnOutils;


procedure TOM_JUEvenement.OnArgument (stArgument:string);
begin
     inherited;
     // rq : ces zones n'existent pas forcément (ex: YYEVEN_MSG)
     if Ecran.Name <> 'YYAGENDA_FIC' then
     begin

          SetControlVisible('TJEV_CODEDOS', false);
          SetControlVisible('TJEV_CODEOP', false);
          SetControlVisible('TTJEV_CODEDOS', false);
          SetControlVisible('TTJEV_CODEOP', false);

          if sGuidPerDos_f <> '' then
          begin
               if sNoDossier_f = '' then
                  sNoDossier_f := GetDosjuri(sGuidPerDos_f);
               if sNoDossier_f = '&#@' then
                  sNoDossier_f := '';
          end;

          if GetControl('DEST')<>Nil then
             THEdit(GetControl('DEST')).OnExit := DEST_OnExit;
     end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 12/02/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUEvenement.OnNewRecord;
begin
   SetField('JEV_DOMAINEACT', '');
   inherited;
end;

procedure TOM_JUEvenement.DEST_OnExit(Sender: TObject);
var tmp, firstdest: String;
begin
  tmp := GetControlText('DEST');
  if tmp<>'' then
    begin
    firstdest := ReadTokenSt(tmp);
    SetField('JEV_EMAIL', firstdest);
    end
  else
    SetField('JEV_EMAIL', '');
end;

Initialization
              RegisterClasses ([TOM_JUEVENEMENT]) ;
end.
