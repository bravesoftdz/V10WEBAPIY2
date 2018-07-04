{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 07/04/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPPLANRUPEDT ()
Mots clefs ... : TOF;CPPLANRUPEDT
*****************************************************************}
Unit CPPLANRUPEDT_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFDEF EAGLCLIENT}
     MainEagl,      // AGLLanceFiche
     eMul,
{$ELSE}
     db,
     dbtables,
     mul,
     FE_main,       // AGLLanceFiche
{$ENDIF}
     Htb97,         // TToolBarbutton97
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOFGRILLEFILTRE,
     UTOF ;

Type
  TOF_CPPLANRUPEDT = Class (TOF_GRILLEFILTRE)

    CC_TYPE : THEdit;
    BDetailPlanRup : TToolBarbutton97;

    // Procédure TOF
    procedure OnArgument ( S : String )                        ; override ;
    procedure OnClose                                          ; override ;
    // Procédure virtuelle
    procedure OnLoadRecord ( vTOBLignes : TOB )                ; override ;
    procedure ParametrageFiche                                 ; override ;
    procedure ParametrageGrille ( vGrille : THGrid )           ; override ;
    procedure OnCreateTOB ( vTOB : TOB )                       ; override ;
    function  GenererConditionPlus                 : String    ; override ;
    function  OnUpdateRecord ( laTOB : TOB )       : Integer   ; override ;
    function  OnAfterUpdateRecord ( laTOB : TOB )  : Integer   ; override ;
    function  OnDeleteRecord ( laTOB : TOB )       : Integer   ; override ;

    // Détail des plans de rupture
    procedure OnClickBDetailPlanRup ( Sender : TObject );

  end ;

procedure PlanRupture(LeCode : String) ;

Implementation

uses Rupture_TOM ; // DetailPlanrupture

////////////////////////////////////////////////////////////////////////////////
procedure PlanRupture(LeCode : String) ;
begin
  AGLLanceFiche('CP', 'CPPLANRUPEDT', '', '', LeCode);
end;

{ TOF_CPPLANRUPEDT }

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPPLANRUPEDT.OnArgument(S: String);
var lStArgument : string;
begin
  inherited;
  lStArgument := S;
  CC_TYPE := THEdit(GetControl('CC_TYPE', True));
  CC_TYPE.Text := ReadTokenSt( lStArgument );

  BDetailPlanRup := TToolBarbutton97(GetControl('BDPLANRUP', True));
  BDetailPlanRup.OnClick := OnClickBDetailPlanRup;

  // Titre de l'écran
  if CC_Type.Text <> '' then
  begin
    case CC_Type.Text[3] of
        'G' : Ecran.Caption := Ecran.Caption + TraduireMemoire('sur les comptes généraux');
        'T' : Ecran.Caption := Ecran.Caption + TraduireMemoire('sur les comptes auxiliaires');
        '1' : Ecran.Caption := Ecran.Caption + TraduireMemoire('sur les sections de l''axe 1');
        '2' : Ecran.Caption := Ecran.Caption + TraduireMemoire('sur les sections de l''axe 2');
        '3' : Ecran.Caption := Ecran.Caption + TraduireMemoire('sur les sections de l''axe 3');
        '4' : Ecran.Caption := Ecran.Caption + TraduireMemoire('sur les sections de l''axe 4');
        '5' : Ecran.Caption := Ecran.Caption + TraduireMemoire('sur les sections de l''axe 5');
        'B' : Ecran.Caption := Ecran.Caption + TraduireMemoire('sur les comptes budgétaires');
    end;
    case CC_Type.Text[3] of
      'G' : Ecran.HelpContext    := 1370000 ;
      'T' : Ecran.HelpContext    := 1380000 ;
      '1'..'5' : Ecran.HelpContext := 1395000 ;
    end;
  end;
  UpdateCaption(Ecran) ;
end;

////////////////////////////////////////////////////////////////////////////////
function TOF_CPPLANRUPEDT.GenererConditionPlus: String;
begin
  Result := '';
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPPLANRUPEDT.OnCreateTOB(vTOB: TOB);
begin
  vTOB.PutValue('CC_TYPE', CC_TYPE.Text ) ;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPPLANRUPEDT.OnLoadRecord(vTOBLignes: TOB);
begin

end;

////////////////////////////////////////////////////////////////////////////////
function TOF_CPPLANRUPEDT.OnAfterUpdateRecord(laTOB: TOB): Integer;
begin
  Result := 0 ;
end;

////////////////////////////////////////////////////////////////////////////////
function TOF_CPPLANRUPEDT.OnDeleteRecord(laTOB: TOB): Integer;
begin
  Result := 0 ;
  try
    ExecuteSQL('DELETE FROM RUPTURE WHERE RU_NATURERUPT = "' + CC_TYPE.Text + '" ' +
               'AND RU_PLANRUPT = "' + laTob.GetValue('CC_CODE') + '"') ;
  except
    on E: Exception do
    begin
      Result := 1;
      PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : OnDeleteRecord');
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
function TOF_CPPLANRUPEDT.OnUpdateRecord(laTOB: TOB): Integer;
begin
  Result := 0 ;
  if Trim(laTOB.GetValue('CC_TYPE')) = '' then
  begin
    Result := 1 ;
    PGIBox('Vous devez renseigner un code.', Ecran.Caption ) ;
  end ;
  if Trim(laTOB.GetValue('CC_LIBELLE')) = '' then
  begin
    Result := 1 ;
    PGIBox('Vous devez renseigner un libellé.', Ecran.Caption ) ;
  end ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPPLANRUPEDT.ParametrageFiche;
begin
  // Param de base
  TitreEcran      := 'Plans de rupture des éditions' ;
  NomTable        := 'CHOIXCOD' ;
  ListeChamps     := 'CC_CODE;CC_LIBELLE' ;
  CodeEtat        := '' ;
  NatureEtat      := '' ;
  ListeColsUniq   := '0';
  ChampTri        := 'CC_CODE' ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPPLANRUPEDT.ParametrageGrille(vGrille: THGrid);
begin
  // Paramétrage de la grille
  vGrille.ColLengths[1] := 3;  // Code 3 char maxi.
  vGrille.ColLengths[2] := 35; // Libellé 35 char maxi.
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPPLANRUPEDT.OnClickBDetailPlanRup(Sender: TObject);
begin
  DetailPlanrupture(CC_TYPE.Text, FListe.Cells[1, FListe.Row], taConsult) ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPPLANRUPEDT.OnClose;
begin
  inherited;
  case CC_Type.Text[3] of
    'G' : AvertirTable('ttRuptGene') ;
    'T' : AvertirTable('ttRuptTiers') ;
    '1' : AvertirTable('ttRuptSect1') ;
    '2' : AvertirTable('ttRuptSect2') ;
    '3' : AvertirTable('ttRuptSect3') ;
    '4' : AvertirTable('ttRuptSect4') ;
    '5' : AvertirTable('ttRuptSect5') ;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
Initialization
  registerclasses ( [ TOF_CPPLANRUPEDT ] ) ;
end.
