{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 24/01/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPVISUDECLATVA ()
Mots clefs ... : TOF;CPVISUDECLATVA
*****************************************************************}
Unit CPVISUDECLATVA_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     fiche,
     Fe_Main,  //AGLLanceFiche
{$else} 
     MaineAGL, //AGLLanceFiche
     eFiche,
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls,
     HTB97, 
     HEnt1, 
     HMsgBox,
     uLibExportTVA, 
     UTOF ;

    procedure LanceVisuDeclaTVA(Arguments : string);

Type
  TOF_CPVISUDECLATVA = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    public
      FListe           : THGrid;
      procedure BTraitementClick(TSender : TObject); 
      procedure BImprimerClick(TSender : TObject);
    private
      maDeclaration    : TDeclaration;
      numDeclaration   : Integer;
      COLNAME          : string;
  end ;

Implementation
uses
{$IFNDEF EAGLCLIENT}
EdtREtat;
{$ELSE}
utileAgl;
{$ENDIF}

procedure LanceVisuDeclaTVA(Arguments : string);
begin
    AGLLanceFiche('CP','CPVISUDECLATVA','','',Arguments);
end;

procedure TOF_CPVISUDECLATVA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPVISUDECLATVA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPVISUDECLATVA.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPVISUDECLATVA.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPVISUDECLATVA.OnArgument (S : String ) ;
var
  x: integer;
  critere: string;
  Arg, Val: string;
begin
  Inherited ;

  // Gestion des arguments
  repeat
    Critere := uppercase(Trim(ReadTokenSt(S)));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        Arg := copy(Critere, 1, x - 1);
        Val := copy(Critere, x + 1, length(Critere));
        if Arg = 'CODECPT' then NumDeclaration := StrToInt(Val);
      end;
    end;
  until Critere = '';

  // Création des objets
  maDeclaration := TDeclaration.Create(NumDeclaration);

  if maDeclaration.isOK then
  begin
     FListe := THGrid(GetControl('FListe'));
     // Renseignement de la fiche
     if maDeclaration.TypeDeclaration = '723' then
     begin
       COLNAME := 'E_CONTREPARTIEAUX;PAYS;T_NIF;CODE_MODIF;MONTANT_CA;TRIMESTRE_MODIF;ANNEE_MODIF'
     end
     else if maDeclaration.TypeDeclaration = '725' then
     begin
        COLNAME := 'E_CONTREPARTIEAUX;T_NIF;MONTANT_CA;MONTANT_TVA';
     end;
     maDeclaration.Traitement.PutGridDetail(FListe,false,false,COLNAME,false);

     if maDeclaration.TypeDeclaration = '723' then
     begin
        // Alignement du compte auxiliaire à gauche :
        FListe.ColAligns[1] := taLeftJustify;
        // Alignement du numero TVA
        FListe.ColAligns[3] := taRightJustify;
        // Alignement du Montant HT
        FListe.ColAligns[5] := taRightJustify;
        FListe.Cells[0,0] := '';
        FListe.Cells[1,0] := TraduireMemoire('Auxiliaire');
        FListe.Cells[2,0] := TraduireMemoire('Pays');
        FListe.Cells[3,0] := TraduireMemoire('N° TVA');
        FListe.Cells[4,0] := TraduireMemoire('Motif');
        FListe.Cells[5,0] := TraduireMemoire('Montant');
        FListe.Cells[6,0] := TraduireMemoire('Trimestre');
        FListe.Cells[7,0] := TraduireMemoire('Annee');
     end
     else if maDeclaration.TypeDeclaration = '725' then
     begin  
        // Alignement du compte auxiliaire à gauche :
        FListe.ColAligns[1] := taLeftJustify;
        // Alignement du numero TVA
        FListe.ColAligns[2] := taRightJustify;
        // Alignement du Montant HT
        FListe.ColAligns[3] := taRightJustify;
        // Alignement du Montant TVA
        FListe.ColAligns[4] := taRightJustify;
        FListe.Cells[0,0] := '';
        FListe.Cells[1,0] := TraduireMemoire('Auxiliaire');
        FListe.Cells[2,0] := TraduireMemoire('N° TVA');
        FListe.Cells[3,0] := TraduireMemoire('Montant HT');
        FListe.Cells[4,0] := TraduireMemoire('Montant TVA');

     end;
     TToolbarButton97(GetControl('BImprimer')).Visible := true;
     // Gestion des evenements
     TToolbarButton97(GetControl('BImprimer')).OnClick := BImprimerClick;
     TToolbarButton97(GetControl('BTraitement')).OnClick := BTraitementClick;
  end
  else
  begin 
     TToolbarButton97(GetControl('BTRAITEMENT')).Visible := false; 
     TToolbarButton97(GetControl('BValider')).Visible := false;
     TToolbarButton97(GetControl('BImprimer')).Visible := false;
     LastError := maDeclaration.ErrorMsg;
     LastErrorMsg := MESS[maDeclaration.ErrorMsg];
  end;
end;

procedure TOF_CPVISUDECLATVA.BImprimerClick;
begin
  if maDeClaration.TypeDeclaration = '723' then
     LanceEtatTob('E','BEL','VIS',maDeclaration.Traitement,True,False,False,TPageControl(GetControl('Pages')),'',Ecran.Caption,False)
  else if maDeClaration.TypeDeclaration = '725' then
     LanceEtatTob('E','BEL','VI2',maDeclaration.Traitement,True,False,False,TPageControl(GetControl('Pages')),'',Ecran.Caption,False);
end;

procedure TOF_CPVISUDECLATVA.BTraitementClick;
var
  Traitement : integer;
begin
  Traitement := mrNo;
  if not(maDeclaration.isValide) then
     Traitement := PGIASK('Etes-vous sûr de vouloir valider cette déclaration ?')  // SDA le 12/12/2007 correction faute ortho
  else
     Traitement := PGIASK('Voulez-vous rééditer cette déclaration ?');
  if (Traitement = mrYes) then
  LanceTraitement(maDeclaration.Numero,not(maDeclaration.isValide));
end;

procedure TOF_CPVISUDECLATVA.OnClose ;
begin
  Inherited ;
  FreeAndNil(maDeclaration);
end ;

procedure TOF_CPVISUDECLATVA.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPVISUDECLATVA.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CPVISUDECLATVA ] ) ; 
end.
