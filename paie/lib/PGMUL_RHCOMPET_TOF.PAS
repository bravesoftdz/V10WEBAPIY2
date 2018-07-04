{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 18/08/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMUL_RHCOMPET ()
Mots clefs ... : TOF;PGMUL_RHCOMPET
*****************************************************************
PT1 24/11/2004 V_60 JL Modifs pour GPAO
}
Unit PGMUL_RHCOMPET_TOF ;

Interface

Uses Classes,
     HQRy,
{$IFDEF AFFAIRE}      
     Hent1,
{$ENDIF AFFAIRE}
{$IFNDEF EAGLCLIENT}
     db,
     mul, HDB,
     FE_main,
{$else}
     eMul,
     uTob,
     Maineagl,
{$ENDIF}
     sysutils,  
     HCtrls,
     HMsgBox,
     UTOF ;

Type
  TOF_PGMUL_RHCOMPET = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    {$IFDEF PAIEGRH}
    private
      Arg : String;
      procedure ValidationSelection(Sender : TObject);
    {$ENDIF}
  end ;

function RHLanceFichePGMUL_RHCOMPET_TOF(Lequel, Argument: string): string;
Implementation

function RHLanceFichePGMUL_RHCOMPET_TOF(Lequel, Argument: string): string;
begin
  result := AGLLanceFiche('PAY','PGMUL_RHCOMPET', '', Lequel, Argument);
end;

procedure TOF_PGMUL_RHCOMPET.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGMUL_RHCOMPET.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGMUL_RHCOMPET.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGMUL_RHCOMPET.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGMUL_RHCOMPET.OnArgument (S : String ) ;
var
  i   : Integer;
  st  : string;
  FieldName : String;
Begin
  Inherited;
{-- Libellés libres ------------------------------------------------------------------------------}
  for i := 1 to 5 do
  begin
    FieldName := 'PCO_TABLELIBRERH'+IntToStr(I);
    if assigned(getControl(FieldName)) then
    begin
      St := RechDom('GCZONELIBRE','RH'+IntToStr(I),FALSE);
      SetControlText('T'+FieldName, St);
      SetControlProperty('T' + FieldName, 'VISIBLE', GetControlText('T' + FieldName) <> '.-');
      SetControlProperty(FieldName, 'VISIBLE', GetControlText('T' + FieldName) <> '.-');
    end;
  end;
{-------------------------------------------------------------------------------------------------}
{$IFDEF PAIEGRH}
  Arg := ReadTokenPipe(S,';');
  If Arg = 'SIMULATIONPOSTE' then
  begin
        TFMul(Ecran).BOuvrir.OnClick := ValidationSelection;
        SetControlVisible('BINSERT',False);
        SetControlVisible('BSELECTALL',True);
        SetControlProperty ('FListe','Multiselection', True);
  end;
  {$ENDIF}
  {$IFDEF SAV}
  Ecran.helpContext := 114000200;
  {$ENDIF SAV}
  {$IFDEF AFFAIRE}
  ecran.caption := traduireMemoire('Compétences');
  UpdateCaption(Ecran);
  {$ENDIF AFFAIRE}  
End;
 
procedure TOF_PGMUL_RHCOMPET.OnClose ;
{$IFDEF PAIEGRH}
var    {$IFNDEF EAGLCLIENT}
       Grille : THDBGrid;
       {$ELSE}
       Grille : THGrid;
        {$ENDIF}
        Q_Mul : THQuery;
        i : Integer;
        LesCompetences : String;
{$ENDIF}
begin
  Inherited ;
  {$IFDEF PAIEGRH}
     If Arg = 'SIMULATIONPOSTE' then
     begin
     {$IFNDEF EAGLCLIENT}
        Grille  :=  THDBGrid(GetControl('FLISTE'));
        {$ELSE}
        Grille  :=  THGrid(GetControl('FLISTE'));
        {$ENDIF}
        Q_Mul := THQuery(Ecran.FindComponent('Q'));
        If (Grille=nil) then Exit;
        LesCompetences := '';
        if (Grille.NbSelected=0) and (not Grille.AllSelected) then
        begin
                MessageAlerte('Aucun élément sélectionné');
                exit;
        end;
        if ((Grille.nbSelected)>0) AND (not Grille.AllSelected ) then
        begin
                for i := 0 to Grille.NbSelected-1 do
                begin
                        Grille.GotoLeBookmark(i);
                        If LesCompetences = '' then LesCompetences := Q_Mul.FindField('PCO_COMPETENCE').AsString
                        else LesCompetences := LesCompetences + ';'+ Q_Mul.FindField('PCO_COMPETENCE').AsString;
                end;
        Grille.ClearSelected;
        end;
        If (Grille.AllSelected=TRUE) then
        begin
                Q_Mul.First;
                while Not Q_Mul.EOF do
                begin
                        If LesCompetences = '' then LesCompetences := Q_Mul.FindField('PCO_COMPETENCE').AsString
                        else LesCompetences := LesCompetences + ';'+ Q_Mul.FindField('PCO_COMPETENCE').AsString;
                        Q_Mul.Next;
                end;
                Grille.AllSelected := False;
        end;
        TFMul(Ecran).Retour:=LesCompetences;
     end;
     {$ENDIF}
end ;

procedure TOF_PGMUL_RHCOMPET.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGMUL_RHCOMPET.OnCancel () ;
begin
  Inherited ;
end ;

{$IFDEF PAIEGRH}
procedure TOF_PGMUL_RHCOMPET.ValidationSelection(Sender : TObject);
begin
  TFMul(Ecran).BAnnulerClick(TFMul(Ecran).BAnnuler);
end;
{$ENDIF}
 
Initialization
  registerclasses ( [ TOF_PGMUL_RHCOMPET ] ) ;
end.
