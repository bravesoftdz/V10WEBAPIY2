{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/12/2002
Modifié le ... :   /  /
Description .. : Source TOF du TobViewer de visualisation du suivi des pièces.
Mots clefs ... : TOF;
*****************************************************************}
Unit UTOFSUIVIPIECE ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_MAIN,       // Pour AGLLANCEFICHE
{$ELSE}
      MainEagl,     // Pour AglLanceGiche en eAgl
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     stat,         // pour le tobViewer
     uTob,         // Pour ShowMeTheTob
     EntGC,uEntCommun,UtilTOBPiece ;       // Pour R_CleDoc

Type

  TOF_SUIVIPIECE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

    Private
      TobAplat : Tob;
      CDRef : R_CleDoc;
      nPiece : Integer;
      TypeSortie, sTiers, sTypeTiers : String;
  end ;

Implementation

Uses UVOIRSUIVIPIECE,
{V500_007 Début}
  FactTOB,
{V500_007 Fin}
  {$IFNDEF EAGLCLIENT}
    edtRetat;
  {$ELSE}
    utilEagl;
  {$ENDIF}


{ TobViewer ----------------------------------------------------------------------------------------------}

procedure TOF_SUIVIPIECE.OnArgument (S : String ) ;
{V500_007 Début}
var
  sWhere : string;
  sDatepiece : string;
{V500_007 Fin}
begin
  Inherited ;
  NPIECE := ValeurI(ReadTokenSt(S));
  CDRef.Naturepiece := ReadTokenSt(S);
  CDRef.Souche := ReadTokenSt(S);
  CDRef.NumeroPiece := ValeurI(ReadTokenSt(S));
  CDRef.Indice := ValeurI(ReadTokenSt(S));
  TypeSortie := ReadTokenSt(S);
  sTiers := ReadTokenSt(S);
  sTypeTiers := ReadTokenSt(S);
  //TobAplat.free;
  If sTypeTiers = '' then
  Begin
    SetControlVisible('NATUREPIECEG',False);
    SetControlVisible('HLNATURE',False);
{V500_007 Début}
    // On viens depuis une pièce => on initialise la fourchette de date à la date de la pièce.
    sWhere := WherePiece(CDRef, ttdPiece, False);
    sDatePiece := GetColonneSQL ('PIECE', 'GP_DATECREATION', sWhere );
    SetControlText('MINDATEPIECE',sDatePiece);
    SetControlText('MAXDATEPIECE',sDatePiece);
{V500_007 Fin}
  End;
end ;

procedure TOF_SUIVIPIECE.OnNew ;
begin
   Inherited ;
end ;

procedure TOF_SUIVIPIECE.OnDelete ;
begin
   Inherited ;
end ;

procedure TOF_SUIVIPIECE.OnUpdate ;
Var
  sWhere : String;
  sListe : String;

  Function MakeListe(sMultival : STring) : String;
  var sItem : STring;
    sResult : String;
  Begin
    sItem := ReadTokenSt(sMultival);
    If sItem <> '' then
      sResult := '"'+sItem+'"';
    While sMultival <> '' do
    Begin
      sItem := ReadTokenSt(sMultival);
      sResult := sResult +',"'+sItem+'"';
    End;
    Result := sResult;
  End;

begin
  if THRadioGroup(GetControl('TYPESORTIE')).ItemIndex = 1 then
    TypeSortie := 'TV'
  Else
    TypeSortie := 'E';
  If GetControltext('MINDATEPIECE') <> '' then
    sWhere := ' AND GL_DATEPIECE >= "'+USdatetime(StrToDate(thedit(GetControl('MINDATEPIECE')).Text))+'"' // Where complémentaire pour le chargement des pièces de base.
  Else
    sWhere := '';
  If GetControltext('MAXDATEPIECE') <> '' then
    sWhere := sWhere + ' AND GL_DATEPIECE <= "'+USdatetime(StrToDate(thedit(GetControl('MAXDATEPIECE')).Text))+'"'; // Where complémentaire pour le chargement des pièces de base.
  If (thMultiValComboBox(GetControl('NATUREPIECEG')).text <> '') and (thMultiValComboBox(GetControl('NATUREPIECEG')).text <> '<<Tous>>') then
  Begin
    sListe := MakeListe(GetControltext('NATUREPIECEG'));
    sWhere := sWhere + ' AND GL_NATUREPIECEG IN ('+sListe+')';
  End;
  If (sTypeTiers = 'CL') and (sTiers <> '') then
    TobAplat := Genere_Tob_Client(sTiers,TypeSortie,sWhere)
  Else
    If (sTypeTiers = 'FO') and (sTiers <> '') then
      TobAplat := Genere_Tob_Fournisseur(sTiers,TypeSortie,sWhere)
    Else
      TobAplat := Genere_Tob(nPiece,CDRef,TypeSortie,sWhere);
  if TobAplat <> nil then
  begin
    if THRadioGroup(GetControl('TYPESORTIE')).ItemIndex = 1 then
    Begin
      TFStat(Ecran).LaTOB :=  TObAplat;
      TFStat(Ecran).ColNames := ';_ORIG;_RUPTURE;_IORDRE;_LIBELLE;GL_NUMERO;GL_NATUREPIECEG;GL_INDICEG;GL_NUMLIGNE;'+
                                'GL_CODEARTICLE;GL_PIECEPRECEDENTE;GL_LIBELLE;GL_QTERESTE;GL_MTRESTE;GL_QTESTOCK;GL_DATEPIECE;'+
                                'GL_QUALIFQTESTO;GL_PUHTDEV;GL_DEVISE;GL_MONTANTHTDEV;GA_LIBELLE;GDI_DIM1;'+
                                'GDI_DIM2;GDI_DIM3;GDI_DIM4;GDI_DIM5;D_SYMBOLE;';  //Colonnes visibles
      Inherited ;
    end
    Else
    begin
      if TobAplat <> nil then
      begin
        LanceEtatTob('E', 'WD0', 'DV0', TobAplat, True, False, False, nil, '', '', False);
        TFStat(Ecran).Close;
      end;
    end;
  end
  else
  begin
    HShowMessage('0;'+Traduirememoire('Synthèse des pièces;Aucune ligne ne correspond à ces critères')+';W;O;O;O;','','');
  end;
end ;

procedure TOF_SUIVIPIECE.OnLoad ;
begin
   Inherited ;
end ;


procedure TOF_SUIVIPIECE.OnClose ;
begin
  Inherited ;
end ;
{ Fin TobViewer ------------------------------------------------------------------------------------------}

Initialization
  registerclasses ( [ TOF_SUIVIPIECE ] ) ;
end.
