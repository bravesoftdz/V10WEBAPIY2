unit CPObjetGen;

interface

uses
  Classes, CPTypeCons;

type

  TZHalleyUser = class(TPersistent)
  private
    FGrpMontantMin     : double ;
    FGrpMontantMax     : double ;
    FJalAutorises      : string ;
    FDossier           : string ;
  public

    constructor Create ( vDossier : String = '' ) ; virtual ;
    procedure   ChargeHalleyUser ;

    property GrpMontantMin    : double read FGrpMontantMin ;
    property GrpMontantMax    : double read FGrpMontantMax ;
    property JalAutorises     : string read FJalAutorises ;
  end;



  TZInfoCpta = class
  private
    FSaisieTranche  : array [1..MaxAxe] of Boolean ;
    FCpta           : array [fbAxe1..fbAux] of TInfoCpta ;
    FLgTableLibre   : TLgTableLibre ;
    FSousPlanAxe    : TSousPlanAxe ;
    FDossier        : string ;
    fMPACC           : TStrings ;

    function        GetCpta ( Value : TFichierBase ) : TInfoCpta ;
    function        GetLgTableLibre ( i , j : Integer) : Integer ;
    function        GetSousPlanAxe ( fb : TFichierBase ; vNumPlan : integer ) : TSousPlan ;
  public

    constructor Create( vDossier : String = '' ) ;
    destructor  Destroy; override ;

    procedure   ChargeSousPlanAxe ;
    procedure   ChargeLgDossier ;
    function    local : boolean ;
    Procedure   ChargeMPACC ;

    property Cpta        [ Value : TFichierBase]                    : TInfoCpta read GetCpta ;
    property LgTableLibre[ i , j : Integer]                         : integer   read GetLgTableLibre ;
    property SousPlanAxe [vAxe : TFichierBase ; vNumPlan : integer] : TSousPlan read GetSousPlanAxe ;
    property MPACC : TStrings read fMPACC write fMPACC;
  end;

  TZCatBud = class
  private
    FCatBud : array[1..MaxCatBud] of TUneCatBud;

    function GetCatBud(Value : Integer) : TUneCatBud;
  public
    constructor Create(vDossier : string = '');
    property    CatBud[aCatBud : Integer] : TUneCatBud read GetCatBud;
  end;

  TCompensation = class
  public
    class function IsCompensation: Boolean; {Retourne True si un plan de compansation est utilisé}
    class function IsPlanCorresp(NumPlan: Integer): Boolean;  {Permet de savoir si un plan de correspondance est utilisé pour la compensation}
    class function GetSQLCorrespCompte(NatureCpt, Compte: String): String; {Nature compte C,1: Client - F,2: Fournisseur}
    class function IsNatureClient(Nature: String): Boolean;
    class function IsNatureFournisseur(Nature: String): Boolean;
    class function GetNomPlanCorrespondance(NatureCpt, Compte: String): String; {Retourne le plan de correspondance associé au compte}
    class function GetNumPlanCorrespondance: Integer;
    class function GetChampPlan: String;
    class function GetJalCompensation: String;
  end;

implementation

uses
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF EAGLCLIENT}
  UTob, HCtrls, HEnt1, CPProcMetier, CPProcGen, SysUtils, ParamSoc;


{---------------------------------------------------------------------------------------}
{------------------------------    TZHalleyUser   --------------------------------------}
{---------------------------------------------------------------------------------------}

{---------------------------------------------------------------------------------------}
constructor TZHalleyUser.Create ( vDossier : string = '' ) ;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TTW}
  cWA.MessagesAuClient('COMSX.IMPORT','','TZHalleyUser.Create') ;
  {$ENDIF}

  FDossier := vDossier;

  ChargeHalleyUser;
end;

{---------------------------------------------------------------------------------------}
procedure TZHalleyUser.ChargeHalleyUser ;
{---------------------------------------------------------------------------------------}
var
  lQG : TQuery ;
begin
  {$IFDEF TTW}
  cWA.MessagesAuClient('COMSX.IMPORT','','TZHalleyUser.ChargeHalleyUser') ;
  {$ENDIF}

  FGrpMontantMin := 0 ;
  FGrpMontantMax :=999999999999.0 ;
  FJalAutorises  := '' ;

  lQG := OpenSQL('SELECT UG_NIVEAUACCES, UG_NUMERO, UG_LIBELLE, UG_CONFIDENTIEL, UG_MONTANTMIN, UG_MONTANTMAX, UG_JALAUTORISES, UG_LANGUE, UG_PERSO FROM USERGRP WHERE UG_GROUPE="' + V_PGI.Groupe + '"', TRUE , -1, 'USERGRP', false ,FDossier) ;

  if not lQG.EOF then
  begin

    FGrpMontantMin := lQG.FindField('UG_MONTANTMIN').AsFloat ;
    FGrpMontantMax := lQG.FindField('UG_MONTANTMAX').AsFloat ;
    FJalAutorises  := lQG.FindField('UG_JALAUTORISES').AsString ;

    // CA - 18/07/2007 - FQ 20027 - Cas du << Tous >> enregistré dans la base
    if (Pos('<<',FJalAutorises)>0) then FJalAutorises:= '';

    if FJalAutorises = '-' then FJalAutorises := '' ; // 13569

    if FJalAutorises <> '' then
    begin
      if FJalAutorises[1] <> ';' then FJalAutorises := ';' + FJalAutorises ;
      if FJalAutorises[Length(FJalAutorises)] <> ';' then FJalAutorises := FJalAutorises + ';' ;
    end ;

    if ( FGrpMontantMax <= FGrpMontantMin) then
    begin
      FGrpMontantMin := 0 ;
      FGrpMontantMax :=999999999999.0 ;
    end ;

  end ;

  Ferme(lQG) ;

end;


{---------------------------------------------------------------------------------------}
{-------------------------------     TZInfoCpta   --------------------------------------}
{---------------------------------------------------------------------------------------}

{---------------------------------------------------------------------------------------}
procedure TZInfoCpta.ChargeMPACC ;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery ;
  i : integer ;
  St : string ;
begin
  MPACC.Clear ;
  Q:=OpenSQl('SELECT MP_MODEPAIE, MP_CODEACCEPT, MP_CATEGORIE, MP_LETTRECHEQUE, MP_LETTRETRAITE FROM MODEPAIE',True) ;
  while not Q.EOF do begin
    St:='' ;
    for i:=0 to 4 do
      St:=St+Q.Fields[i].AsString+';' ;
    MPACC.Add(St) ;
    Q.Next ;
  end;
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
constructor TZInfoCpta.Create ( vDossier : String = '' ) ;
{---------------------------------------------------------------------------------------}
var
 StB            : string ;
 QLoc , QLoc1,Q : TQuery ;
 NumAxe         : integer ;
 fb             : TFichierBase ;
 CO,CL          : String ;
 Lg,i,j         : integer;
begin
  {$IFDEF TTW}
  cWA.MessagesAuClient('COMSX.IMPORT','','TZInfoCpta.Create') ;
  {$ENDIF}
  MPACC := TStringList.Create;

  FDossier := vDossier ;

  ChargeLgDossier ;

  QLoc:=OpenSQL('Select * From STRUCRSE Where SS_CONTROLE="X" ',True,-1, 'STRUCRSE',false , FDossier) ;
  while not QLoc.Eof do begin
    fb:=AxeToFb(QLoc.FindField('SS_AXE').AsString) ;
    FCpta[fb].UneStruc.CodeSp:=QLoc.FindField('SS_SOUSSECTION').AsString ;
    FCpta[fb].UneStruc.Debu:=QLoc.FindField('SS_DEBUT').AsInteger ;
    FCpta[fb].UneStruc.Long:=QLoc.FindField('SS_LONGUEUR').AsInteger ;
    QLoc1:=OpenSql('Select PS_CODE From SSSTRUCR Where PS_AXE="'+QLoc.FindField('SS_AXE').AsString+'" And '+
                   'PS_SOUSSECTION="'+QLoc.FindField('SS_SOUSSECTION').AsString+'" ',True,-1,'',false,FDossier) ;
    FCpta[fb].UneStruc.QuelCarac:=QLoc1.Fields[0].AsString ; Ferme(QLoc1) ;
    QLoc.Next ;
  end;
  Ferme(QLoc) ;

  NumAxe:=0 ;
  Q:=OpenSQL('SELECT * FROM AXE ORDER BY X_AXE',TRUE,-1,'AXE',false,FDossier) ;
  while not Q.EOF do begin
    Inc(NumAxe) ; fb:=AxeToFb(Q.FindField('X_AXE').AsString) ;
    FCpta[fb].Lg:=Q.FindField('X_LONGSECTION').AsInteger ;
    Stb:=Q.FindField('X_BOURREANA').AsString ;
    if Stb <> '' then
      FCpta[fb].Cb:=Stb[1]
    else
      FCpta[fb].Cb:='0' ;
    FSaisieTranche[NumAxe]:=(Q.FindField('X_SAISIETRANCHE').AsString='X') ;
    FCpta[fb].Chantier:=Q.FindField('X_CHANTIER').AsString='X' ;
    FCpta[fb].Structure:=Q.FindField('X_STRUCTURE').AsString='X' ;
    FCpta[fb].Attente:=Q.FindField('X_SECTIONATTENTE').AsString ;
    FCpta[fb].AxGenAttente:=Q.FindField('X_GENEATTENTE').AsString ;
    Q.Next ;
  end ;
  Ferme(Q) ;

  // Tables libres
  Fillchar(FLgTableLibre, SizeOf(FLgTableLibre),#0) ;
  Q:=OpenSQL('SELECT CC_CODE,CC_LIBRE,CC_TYPE FROM CHOIXCOD WHERE CC_TYPE="NAT" ORDER BY CC_TYPE,CC_CODE',TRUE,-1,'CHOIXCOD',false,FDossier) ;
  while not q.EOF do begin
    CO:=Q.Fields[0].AsString ;
    CL:=Q.Fields[1].AsString ;
    if OkChar(CL) then begin
      Lg:=StrToInt(CL) ;
      case CO[1] of
        'G' : i:=1 ; 'T' : i:=2 ; 'S' : i:=3 ; 'B' : i:=4 ;
        'D' : i:=5 ; 'E' : i:=6 ; 'A' : i:=7 ; 'U' : i:=8 ;
        'I' : i:=9 ;
      else
        i:=0 ;
      end ;

      if i>0 then begin
        j:=StrToInt(Copy(CO,2,2))+1 ;
        FLgTableLibre[i,j]:=Lg ;
      end ;
    end ;
    Q.Next ;
  end ;
  Ferme(Q) ;

  ChargeSousPlanAxe ;
  ChargeMPACC ;
end;

{---------------------------------------------------------------------------------------}
destructor TZInfoCpta.Destroy;
{---------------------------------------------------------------------------------------}
var
 fb : TFichierBase ;
 i  : integer ;
begin

  for fb := fbAxe1 to fbAxe5 do
    for i := 1 To MaxSousPlan do
      if FSousPlanAxe[fb,i].ListeSP <> nil then
      begin
        FSousPlanAxe[fb,i].ListeSP.Clear ;
        FSousPlanAxe[fb,i].ListeSP.Free ;
      end ;

  MPAcc.Free;
  inherited ;
end;

{---------------------------------------------------------------------------------------}
procedure TZInfoCpta.ChargeLgDossier ;
{---------------------------------------------------------------------------------------}
var
 Stb : string ;
begin
  {$IFDEF TTW}
  cWA.MessagesAuClient('COMSX.IMPORT','','TZInfoCpta.ChargeLgDossier') ;
  {$ENDIF}

  if local then begin
    FCpta[fbGene].Lg      := GetParamSocSecur('SO_LGCPTEGEN', 0 ) ;
    FCpta[fbGene].Attente := GetParamSocSecur('SO_GENATTEND', '') ;
    Stb                   := GetParamSocSecur('SO_BOURREGEN', '') ;
  end
  else begin
    FCpta[fbGene].Lg      := GetParamSocDossierSecur('SO_LGCPTEGEN', 0 , FDossier) ;
    FCpta[fbGene].Attente := GetParamSocDossierSecur('SO_GENATTEND', '', FDossier) ;
    Stb                   := GetParamSocDossierSecur('SO_BOURREGEN', '', FDossier) ;
  end ;

  if Stb <> '' then
    FCpta[fbGene].Cb := Stb[1]
  else
    FCpta[fbGene].Cb := '0' ;

  if local then begin
    FCpta[fbAux].Lg       := GetParamSocSecur('SO_LGCPTEAUX', 0) ;
    Stb                   := GetParamSocSecur('SO_BOURREAUX', '') ;
  end
  else begin
    FCpta[fbAux].Lg       := GetParamSocDossierSecur('SO_LGCPTEAUX', 0,  FDossier) ;
    Stb                   := GetParamSocDossierSecur('SO_BOURREAUX', '', FDossier) ;
  end ;

  if Stb <> '' then
    FCpta[fbAux].Cb := Stb[1]
  else
    FCpta[fbAux].Cb:='0' ;

end ;


{---------------------------------------------------------------------------------------}
procedure TZInfoCpta.ChargeSousPlanAxe ;
{---------------------------------------------------------------------------------------}
var
 fb                : TFichierBase ;
 Q,Q1              : TQuery ;
 QAxe,QCode,QLib   : String ;
 QDeb,QLong        : Integer ;
 TabI              : array[fbAxe1..fbAxe5] Of Integer ;
 i                 : integer ;
begin
  {$IFDEF TTW}
  cWA.MessagesAuClient('COMSX.IMPORT','','TZInfoCpta.ChargeSousPlanAxe') ;
  {$ENDIF}

  for fb := fbAxe1 to fbAxe5 do
    for i := 1 To MaxSousPlan do
      if FSousPlanAxe[fb,i].ListeSP <> nil then
      begin
        FSousPlanAxe[fb,i].ListeSP.Clear ;
        FSousPlanAxe[fb,i].ListeSP.Free ;
      end ;

  Fillchar(FSousPlanAxe,SizeOf(FSousPlanAxe),#0) ;
  Fillchar(TabI,SizeOf(TabI),#0) ;

  Q := OpenSQL('Select * From STRUCRSE Order By SS_AXE, SS_DEBUT',True,-1,'STRUCRSE',false,FDossier) ;
  while not Q.Eof Do
  begin
    QAxe   := Q.FindField('SS_AXE').AsString ;
    QCode  := Q.FindField('SS_SOUSSECTION').AsString ;
    QLib   := Q.FindField('SS_LIBELLE').AsString ;
    QDeb   := Q.FindField('SS_DEBUT').AsInteger ;
    QLong  := Q.FindField('SS_LONGUEUR').AsInteger ;
    fb     :=AxeTofb(QAxe) ;
    Inc(TabI[fb]) ;

    if tabI[fb] <= MaxSousPlan Then
    begin
      FSousPlanAxe[fb,TabI[fb]].Code       := QCode ;
      FSousPlanAxe[fb,TabI[fb]].Lib        := QLib ;
      FSousPlanAxe[fb,TabI[fb]].Debut      := QDeb ;
      FSousPlanAxe[fb,TabI[fb]].Longueur   := QLong ;
      FSousPlanAxe[fb,TabI[fb]].ListeSP    := HTStringList.Create ;
      Q1 := OpenSQL('Select * from SSSTRUCR WHERE PS_AXE="' + QAxe + '" AND PS_SOUSSECTION="' + QCode + '" ORDER BY PS_CODE' , true , -1 ,'SSSTRUCR',false,FDossier) ;
      while not Q1.Eof Do
      begin
        FSousPlanAxe[fb,TabI[fb]].ListeSP.Add(Q1.FindField('PS_CODE').ASString + ';' + Q1.FindField('PS_LIBELLE').ASString + ';' ) ;
        Q1.Next ;
      end ;
      Ferme(Q1) ;
    end ;
    Q.Next ;
  end ;
  Ferme(Q) ;
end ;

{---------------------------------------------------------------------------------------}
function TZInfoCpta.GetCpta(Value: TFichierBase) : TInfoCpta;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TTW}
  cWA.MessagesAuClient('COMSX.IMPORT','','TZInfoCpta.GetCpta') ;
  {$ENDIF}
  Result := FCpta[Value];
end;

{---------------------------------------------------------------------------------------}
function TZInfoCpta.GetLgTableLibre ( i , j : Integer) : Integer;
{---------------------------------------------------------------------------------------}
begin
  result := FLgTableLibre[i, j];
end;

{---------------------------------------------------------------------------------------}
function TZInfoCpta.GetSousPlanAxe( fb : TFichierBase ; vNumPlan : integer ) : TSousPlan;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TTW}
  cWA.MessagesAuClient('COMSX.IMPORT','','TZInfoCpta.GetSousPlanAxe') ;
  {$ENDIF}
  Result := FSousPlanAxe[fb,vNumPlan] ;
end;

{e FP 21/02/2006}
{---------------------------------------------------------------------------------------}
function TZInfoCpta.local: boolean;
{---------------------------------------------------------------------------------------}
begin
  result := (FDossier = '') or ( FDossier = V_PGI.SchemaName ) ;
end;

{---------------------------------------------------------------------------------------}
{------------------------------   TCompensation   --------------------------------------}
{---------------------------------------------------------------------------------------}

{---------------------------------------------------------------------------------------}
class function TCompensation.GetChampPlan: String;
{---------------------------------------------------------------------------------------}
begin
  Result := '';
  if IsCompensation then
    Result := 'T_CORRESP'+IntToStr(GetNumPlanCorrespondance);
end;

{---------------------------------------------------------------------------------------}
class function TCompensation.GetJalCompensation: String;
{---------------------------------------------------------------------------------------}
begin
  Result := GetParamSocSecur('SO_CPJALCOMPENSATION', '');
end;

{---------------------------------------------------------------------------------------}
class function TCompensation.GetNomPlanCorrespondance(NatureCpt, Compte: String): String;
{---------------------------------------------------------------------------------------}
var
  NumPlan: Integer;
  SQL:     String;
  Q:       TQuery;
begin
  Result := '';
  if not IsCompensation then
    Exit;
  if IsPlanCorresp(1) then
    NumPlan := 1
  else
    NumPlan := 2;

  SQL := 'SELECT CR_CORRESP'+
         ' FROM CORRESP'+
         ' WHERE CR_TYPE="AU'+IntToStr(NumPlan)+'"'+
         '   AND '+GetSQLCorrespCompte(NatureCpt, Compte);
  Q := OpenSQL(SQL, True);
  if not Q.EOF then
    Result := Q.FindField('CR_CORRESP').AsString;
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
class function TCompensation.GetNumPlanCorrespondance: Integer;
{---------------------------------------------------------------------------------------}
begin
  Result := 0;
  if IsPlanCorresp(1) then
    Result := 1
  else if IsPlanCorresp(2) then
    Result := 2;
end;

{---------------------------------------------------------------------------------------}
class function TCompensation.GetSQLCorrespCompte(NatureCpt, Compte: String): String;
{---------------------------------------------------------------------------------------}
begin
  if IsNatureClient(NatureCpt) then
    Result := ' CR_LIBRETEXTE1="'+Compte+'"'
  else if IsNatureFournisseur(NatureCpt) then {Fournisseur}
    Result := ' CR_LIBRETEXTE2="'+Compte+'"'
  else
    Result := ' 1<>1';
end;

{---------------------------------------------------------------------------------------}
class function TCompensation.IsCompensation: Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := IsPlanCorresp(1) or IsPlanCorresp(2);
end;

{---------------------------------------------------------------------------------------}
class function TCompensation.IsNatureClient(Nature: String): Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := (Nature='CLI') or (Nature='AUD');
end;

{---------------------------------------------------------------------------------------}
class function TCompensation.IsNatureFournisseur(Nature: String): Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := (Nature='FOU') or (Nature='AUC');
end;

{---------------------------------------------------------------------------------------}
class function TCompensation.IsPlanCorresp(NumPlan: Integer): Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := GetParamSocSecur('SO_CPGESTCOMPENSATION', False);
  if Result then
    Result := IntToStr(NumPlan) = GetParamSocSecur('SO_CPPLANCOMPENSATION', '');
end;

{---------------------------------------------------------------------------------------}
{--------------------------------    TZCatBud     --------------------------------------}
{---------------------------------------------------------------------------------------}

{---------------------------------------------------------------------------------------}
constructor TZCatBud.Create(vDossier : string = '');
{---------------------------------------------------------------------------------------}
var
  fb : TFichierBase;
  Q  : TQuery;
  StCode,
  StLib,
  StLibre,
  StS, StJ,
  St      : string;
  i, k, l : Integer;
begin
  {$IFDEF TTW}
   cWA.MessagesAuClient('COMSX.IMPORT','','TZCatBub.Create') ;
  {$ENDIF}
  Fillchar(FCatBud, SizeOf(FCatBud), #0);
  Q := OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="CJB"', TRUE, -1, 'CHOIXCOD', False , vDossier);
  k := 1;
  while (not Q.Eof) and (k <= MaxCatBud) do begin
    StCode := Q.FindField('CC_CODE'   ).AsString;
    StLib  := Q.FindField('CC_LIBELLE').AsString;
    StLibre:= Q.FindField('CC_LIBRE'  ).AsString;
    fb     := AxeTofb(Q.FindField('CC_ABREGE').AsString);
    i := Pos('/', StLibre);

    if i > 0 then begin
      StS := Copy(StLibre, 1, i - 1);
      StJ := Copy(StLibre, i + 1, Length(StLibre) - i);
      FCatBud[k].fb   := fb;
      FCatBud[k].Code := StCode;
      FCatBud[k].Lib  :=StLib;

      l := 1;
      while StS <> '' do begin
        St := ReadTokenSt(StS);
        if St <> '' then begin
          FCatBud[k].SurSect[l] := St;
          Inc(l);
        end;
      end;

      l := 1;
      while StJ <> '' do begin
       St := ReadTokenSt(StJ);
       if St <> '' then begin
         FCatBud[k].SurJal[l] := St;
         Inc(l);
       end;
     end;
     Inc(k);
    end;
    Q.Next;
  end;
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
function TZCatBud.GetCatBud(Value : Integer) : TUneCatBud;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TTW}
   cWA.MessagesAuClient('COMSX.IMPORT','','TZCatBub.GetCatBud') ;
  {$ENDIF}
   Result := FCatBud[Value];
end;


end.

