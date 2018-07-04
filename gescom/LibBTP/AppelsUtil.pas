unit AppelsUtil;

interface
uses
    Classes,
    HEnt1,
    Paramsoc,
    ComCtrls,
    sysutils,
    HCtrls,
    EntGC,
    hmsgbox,
    LookUp,
    UTob,
    FactUtil,
    AGLInit,
    Controls,
    Formule,
    UtilArticle,
    tiersutil,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,HDB,
{$ENDIF}
     DicoBTP,
     UtilPGI,
     HTB97,
     Forms,
     SaisUtil,
     Graphics,
     AfUtilArticle,
     Ent1,
     UtilRessource ,
     CalcOleGenericBTP,
     BTPUtil,
     FactTOB,
     FactAdresse ;

Const

  NbMaxPartieAppel =3;   // mcd 28/03/03 si modif voir valeur en dur dans source CalcOleGenericAff
  cInClient       = 0; // standard pas de numero de client
  cInClientAlgoe  = 1; // algoe
  cInClientOnyx   = 2; // onyx
  cInClientAmyot  = 3; // Amyot

Var DicoAppel : array[1..2,1..20] of string ;

      TexteMsgPieceAppel: array[1..4] of string 	= (
      {1}        'Code Appel non valide',
      {2}        'Plus d''une pièce associée à l''Appel',
      {3}        'Aucune pièce associée à l''Appel',
      {4}        'Code Tiers non valide'
                 );

Function CompteurPartieAppel( NumPartie : Integer; bProposition,bUpdateCpt : boolean ) : string;
function DechargeCleAppel(Part0,Part1,Part2,Part3,Avenant, CodeTiers : string;
                              FTypeAction:TActionFiche;
                              SaisieAppel,AfficheMsg,bProposition:Boolean;
                              var AviPartErreur:integer; FromContrat : boolean=false):String;

Function RechercheAppel(Aff0,Aff1,Aff2,Aff3,Aff4 : String; EnCreatAff : Boolean):Boolean;
Function InterpreteZoneExercice (TypeExer : string; var Exercice : String; Var AnneeDeb,AnneeFin,Mois : integer; Obligatoire : Boolean) : Boolean;
Function RegroupePartiesAppel (Var Part0, Part1, Part2, Part3,Avenant : String) : string;
function CodeAppelSuivant ( Code : String ; Lng : Integer) : String ;
Function ChargeCodeAppel(Nbpartie : integer;
                         TypeCode, Valeur, Libelle : string;
                         FTypeAction:TActionFiche;
                         SaisieAppel:boolean) : string;

function FormatNoTel(Notel : String) : string;
Function CalculCodeAppel(P0,P1,P2,P3,Avn, Tiers : String;FromContrat : boolean=false) : String;

Procedure CodeAppelDecoupe(CodeApp : string; Var Part0, Part1, Part2, Part3, Avenant : String);

procedure ChargeClefAppel (Part0,Part1,Part2,Part3,Avenant: THCritMaskEDIT;BRechAffaire:TToolBarButton97; FTypeAction:TActionFiche; CodeAff:string;SaisieAffaire:Boolean);
function  GetTypeArticle(Code : string) : string;

{$IFDEF EAGLCLIENT}
Procedure ModifColAppelGrid (Gr : THGrid; TobQ : TOB );
{$ELSE}
Procedure ModifColAppelGrid (Gr : THDBGrid );
{$ENDIF}

procedure PrePrepositionneNewCodeAppel (var Part1,Part2,Part3: string);
procedure FinaliseNewCodeAppel (var Part1,Part2,Part3: string);
function ControleSaisieOK (Part1,Part2,Part3: string) : boolean;


implementation
uses utofAfPiece,AffaireUtil, FactComm;


Function FinaliseCodeAppel(Nopartie : integer) : string;
Begin
  Result := CompteurPartieAppel(Nopartie, false,True);
End;


Function PrepareCodeAppel(Nopartie : integer; TypeCode : string; valeur : string) : string;
Begin

  Result := '';
  if (TypeCode='CPM') then
  begin
    result := CompteurPartieAppel(Nopartie, false,True);;
  end else if (TypeCode ='FIX') then
  begin
    Result := valeur;
  end else if (TypeCode ='DAT') then
  begin
    if length(valeur) = 6 then
    begin
      Result := FormatDateTime('yyyymm', Now);
    end else
    begin
      Result := FormatDateTime('yymm', Now);
    end;
  end;
End;

procedure PrePrepositionneNewCodeAppel (var Part1,Part2,Part3: string);
var TypeCode : string3;
    Valeur,tmp : String;
    i : integer;
begin
  for i:=1 to GetParamSoc('SO_APPCODENBPARTIE') do
  Begin
    case i of
    1: Begin
        TypeCode:=GetParamSoc('SO_APPCO1TYPE');
        Valeur:=GetParamSoc('SO_APPCO1VALEUR');
        Part1 :=PrepareCodeAppel(i,TypeCode,Valeur);
      End;
    2: Begin
        TypeCode:=GetParamSoc('SO_APPCO2TYPE');
        Valeur:=GetParamSoc('SO_APPCO2VALEUR');
        Part2 :=PrepareCodeAppel(i,TypeCode,Valeur);
      End;
    3: Begin
        TypeCode:=GetParamSoc('SO_APPCO3TYPE');
        Valeur:=GetParamSoc('SO_APPCO3VALEUR');
        Part3 :=PrepareCodeAppel(i,TypeCode,Valeur);
      End;
    End;
  End;
end;


procedure FinaliseNewCodeAppel (var Part1,Part2,Part3: string);
var TypeCode : string3;
    Valeur,tmp : String;
    i : integer;
begin
  for i:=1 to GetParamSoc('SO_APPCODENBPARTIE') do
  Begin
    case i of
    1: Begin
        TypeCode:=GetParamSoc('SO_APPCO1TYPE');
        if (TypeCode = 'CPT')then
        begin
          Part1 :=FinaliseCodeAppel(i);
        end;
      End;
    2: Begin
        TypeCode:=GetParamSoc('SO_APPCO2TYPE');
        if (TypeCode = 'CPT')then
        begin
          Part2 :=FinaliseCodeAppel(i);
        end;
      End;
    3: Begin
        TypeCode:=GetParamSoc('SO_APPCO3TYPE');
        if (TypeCode = 'CPT')then
        begin
          Part3 :=FinaliseCodeAppel(i);
        end;
      End;
    End;
  End;
end;

function ControleSaisieOK (Part1,Part2,Part3: string) : boolean;
var TypeCode : string3;
    Valeur,tmp : String;
    i : integer;
begin
  result := true;
  for i:=1 to GetParamSoc('SO_APPCODENBPARTIE') do
  Begin
    case i of
    1: Begin
        TypeCode:=GetParamSoc('SO_APPCO1TYPE');
        if (Pos(TypeCode,'CPT;CPM;FIX;DAT')=0) then
        begin
          if Part1='' then
          begin
            result := false;
            break;
          end;
        end;
      End;
    2: Begin
        TypeCode:=GetParamSoc('SO_APPCO2TYPE');
        if (Pos(TypeCode,'CPT;CPM;FIX;DAT')=0) then
        begin
          if Part2='' then
          begin
            result := false;
            break;
          end;
        end;
      End;
    3: Begin
        TypeCode:=GetParamSoc('SO_APPCO3TYPE');
        if (Pos(TypeCode,'CPT;CPM;FIX;DAT')=0) then
        begin
          if Part3='' then
          begin
            result := false;
            break;
          end;
        end;
      End;
    End;
  End;
end;

function  GetTypeArticle(Code : string) : string;
var QQ: TQuery;
begin
  result := 'MAR';
  QQ := OpenSql('SELECT GA_TYPEARTICLE FROM ARTICLE WHERE GA_ARTICLE="'+Code+'"',True,1,'',true);
  if not QQ.eof then Result := QQ.findField('GA_TYPEARTICLE').AsString;
  ferme (QQ);
end;



procedure ChargeClefAppel (Part0,Part1,Part2,Part3,Avenant: THCritMaskEDIT;BRechAffaire:TToolBarButton97; FTypeAction:TActionFiche; CodeAff:string;SaisieAffaire:Boolean);
var i,lng,NbPartieVisible, PositAvenant : integer;
    TypeCode : string3;
    Valeur,Libelle,CodePartie0,CodePartie1,CodePartie2,CodePartie3,CodeAvenant : String;
    Visible : Boolean;
    Code : THCritMaskEdit;
    OnChangetmp1, OnChangetmp2 : TNotifyEvent;
    HH            : THLabel ;
    FF            : TForm ;
    nbParties : integer;
begin
  lng := 0;  Visible :=False;
  if Not(ctxAffaire in V_PGI.PGIContexte) and Not(ctxGCAFF in V_PGI.PGIContexte) then
    BEGIN
      if Part0  <> Nil then Part0.Visible := False;
      if Part1  <> Nil then Part1.Visible := False;
      if Part2  <> Nil then Part2.Visible := False;
      if Part3  <> Nil then Part3.Visible := False;
      if Avenant<> Nil then Avenant.Visible := False;
      if Part1<>Nil then
        BEGIN
          FF:=TForm(Part1.Owner) ;
          for i:=0 to FF.ComponentCount-1 do if FF.Components[i] is THLabel then
            BEGIN
              HH:=THLabel(FF.Components[i]) ;
              if HH.FocusControl=Part1 then BEGIN HH.Visible:=False ; Break ; END ;
            END ;
         END;
      Exit;
   END;
  OnChangetmp1 := Part1.OnChange;
  OnChangetmp2 := Part2.OnChange;
  Part1.OnChange := Nil;
  Part2.OnChange := Nil;
  Part1.Text :='' ; Part2.Text := ''; Part3.Text := '';

  try
  NbPartieVisible := 0;
  PositAvenant := 0;
  // Gestion des affaires
  if Part0 <> Nil then
    BEGIN
      Part0.DataType:= 'AFSTATUTREDUIT';
      Part0.ElipsisButton:=True;
    END;
  nbParties := GetParamSoc('SO_APPCODENBPARTIE');
  for i:=1 to nbParties do
    Begin
      // recup des valeurs de la partie du code traitée
      case i of
        1: Begin
             Code:=Part1;
             TypeCode:=GetParamSoc('SO_APPCO1TYPE');
             Valeur:=GetParamSoc('SO_APPCO1VALEUR');
             Libelle:=GetParamSoc('SO_APPCO1LIB');
             lng:=GetParamSoc('SO_APPCO1LNG');
             Visible:=True;
           End;
        2: Begin
             Code:=Part2;
             TypeCode:=GetParamSoc('SO_APPCO2TYPE');
             Valeur:=GetParamSoc('SO_APPCO2VALEUR');
             Libelle:=GetParamSoc('SO_APPCO2LIB');
             lng:=GetParamSoc('SO_APPCO2LNG');
             Visible:=GetParamSoc('SO_APPCO2VISIBLE');
           End;
        3: Begin
             Code:=Part3;
             TypeCode:=GetParamSoc('SO_APPCO3TYPE');
             Valeur:=GetParamSoc('SO_APPCO3VALEUR');
             Libelle:=GetParamSoc('SO_APPCO3LIB');
             lng:=GetParamSoc('SO_APPCO3LNG');
             Visible:=GetParamSoc('SO_APPCO3VISIBLE');
           End;
        else
          code := nil;
      End;

      Code.ReadOnly:=True;
      Code.enabled := false;

      If (i <= nbParties) Then
        Begin   // Partie utilisée
          Code.Visible:= Visible;
          Code.Hint:=Libelle; // propriété visible + conseil
          Code.ElipsisButton:=False;
          Code.MaxLength:=lng;
          Code.Width :=(lng * 8)+10; // Calcul longueur des zones
          if ((Typecode='SAI') And (i=2) And (VH_GC.AFFORMATEXER<>'AUC')) then
             Code.Width := Code.Width + 15; // séparateurs en plus
          If (i=2) Then Code.Left:=Part1.Left+Part1.Width+5;
          If (i=3) Then Code.Left:=Part2.Left+Part2.Width+5;

          If (SaisieAffaire) then
              Begin
              if (FTypeAction=taCreat) then Code.ReadOnly:=False else Code.ReadOnly:= True;
              End
          Else // Code affaire depuis une autre saisie
              Begin
              If (FTypeAction <> taConsult) Then Code.ReadOnly:=False else Code.ReadOnly:=True;
              End;


          // spécificitées en fonction du type de zone ...
          if (Typecode= 'FIX')then
              begin
              Code.ReadOnly:=True;
              Code.enabled := false;
              Code.Text:= valeur; //mcd 18/09/01 ne marchait pas ..
              Code.color := ClBtnFace
              end else                                    // Zone fixe
          if (Typecode= 'LIS') then                                       // Liste de saisie
              Begin
              if ((SaisieAffaire) And (FTypeAction = taCreat )) Or
                  (Not(SaisieAffaire) And (FTypeAction <> taConsult) And (Code.ReadOnly<>True)) Then Code.ElipsisButton:=True;
              Code.DataType:='AFFAIREPART'+intToStr(i);   Code.Width:=45;
              End else
          if (Typecode='CPT') then // compteur
              BEGIN
              Code.ReadOnly:=True;
              Code.color := ClBtnFace;
              Code.enabled := false;
              End else
          if (Typecode='CPM') then // compteur modifiable
              BEGIN
              Code.ReadOnly:=false;
              Code.color := clWindow;
              Code.enabled := true;
              End else
          if (TypeCode='DAT') then
              begin
                if (lng = 6) then
                begin
                  Code.Hint :=Code.Hint+ '(AAAAMM)'
                end else
                begin
                  Code.Hint :=Code.Hint+ '(AAMM)'
                end;
              end;
              // mcd déplacer, car longueur change si LIS 26/06/01
          if Visible then
             BEGIN
             PositAvenant := Code.Left +  Code.Width + 5; // Calcul position avenant
             NbPartieVisible := i;
             END;
        End
      Else
        Begin  // RAZ de la partie de structure non utilisée
          Code.Visible:=False;
        End;
    End;

  CodePartie1 := ''; CodePartie2 := ''; CodePartie3 :=''; CodePartie0 := ''; CodeAvenant := '';
  if (codeAff <> '')  then
    begin
			CodeAppelDecoupe (CodeAff,CodePartie0,CodePartie1,CodePartie2,CodePartie3,CodeAvenant);
      Part1.Text :=CodePartie1 ;
      Part2.Text := CodePartie2;
      Part3.Text := CodePartie3;
    end;

  if (Part0<>nil) then Part0.Text :=CodePartie0 ;
  // mcd 18/09/01 dépalcer car non OK si partie FIX Part1.Text :=CodePartie1 ; Part2.Text := CodePartie2; Part3.Text := CodePartie3;

  // positionnement du bouton de recherche affaire
  If ((NbPartieVisible > 0) And (BRechAffaire <> nil)) Then
    Begin
      Case  NbPartieVisible of
          1 : BRechAffaire.Left := Part1.Left + Part1.Width+5;
          2 : BRechAffaire.Left := Part2.Left + Part2.Width+5;
          3 : BRechAffaire.Left := Part3.Left + Part3.Width+5;
      End;
    End;
  // Gestion de la zone avenant
  if Avenant <> Nil then
    BEGIN
      if (VH_GC.CleAppel.GestionAvenant) then
          BEGIN
          Avenant.Visible := true;
          Avenant.Left := PositAvenant; Avenant.Width :=2 * 10 + 5;
          Avenant.Hint := 'Avenant';    Avenant.MaxLength := 2;
          {$IFDEF BTP}
          Avenant.Text := BTPTestcodeAvenant(CodeAvenant,SaisieAffaire);
          {$ELSE}
          Avenant.Text := TestcodeAvenant(CodeAvenant,SaisieAffaire);
          {$ENDIF}
          If (SaisieAffaire) then
              BEGIN
              Avenant.ReadOnly:=True;
              if (FTypeAction=taCreat) then Avenant.color := ClBtnFace
                                       else Avenant.color := ClWindow;
              END;
          END
      else Avenant.Visible := false;
    END;

  finally
    Part1.OnChange:= OnChangetmp1;
    Part2.OnChange:= OnChangetmp2;
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 02/02/2000
Modifié le ... : 02/02/2000
Description .. : Contrôle des 3 parties du code Appel et
                 composition du code Appel concaténé.
Modification.. : Modif PL le 15/06/2000 : ajout de AviPartErreur
                 pour savoir quelle zone sort en erreur
Mots clefs ... : AFFAIRE;CODE AFFAIRE
*****************************************************************}
function DechargeCleAppel(Part0,Part1,Part2,Part3,Avenant,CodeTiers : string;
                          FTypeAction:TActionFiche;
                          SaisieAppel,AfficheMsg,bProposition:Boolean;
                          var AviPartErreur:integer; FromContrat : boolean=false):String;
Var i, AnneeDeb, AnneeFin, Mois : integer;
    Code            : string;
    TypeCode        : String3;
    Exercice        : String;
    ExerObligatoire : Boolean;
    Valeur          : string;
    Libelle         : String;
    StatutReduit    : String;
    CodeAvenant     : String;
    tmp : string;
    //P2              : string;
    //P3              : string;
Begin
  Result:='';
  AviPartErreur:=0;

  // Gestion avenant + statut
  if (Avenant <> '') then CodeAvenant := BTPTestCodeAvenant (Avenant,True)
  									 else CodeAvenant := '00';

  if Part0 <> '' then StatutReduit := Part0
  							 else StatutReduit := 'W';

  if FtypeAction = TaCreat then
  begin
    for i:=1 to GetParamSoc('SO_APPCODENBPARTIE')do
    Begin
      case i of
      1: Begin
          TypeCode:=GetParamSoc('SO_APPCO1TYPE');
          Valeur:=GetParamSoc('SO_APPCO1VALEUR');
          Libelle:=GetParamSoc('SO_APPCO1LIB');
          if not fromContrat then
          begin
          	Part1:=ChargeCodeAppel(i,TypeCode,Valeur,Libelle,FTypeAction,SaisieAppel);
          end else
          begin
          	tmp:=ChargeCodeAppel(i,TypeCode,Valeur,Libelle,FTypeAction,SaisieAppel);
          	if ((TypeCode = 'FIX') and (pos(Tmp,'0; ')=0)) or (TypeCode <>'FIX') then
            begin
              Part1 := Tmp;
            end;
          end;
        End;
      2: Begin
          Code:=Part2;
          TypeCode:=GetParamSoc('SO_APPCO2TYPE');
          Valeur:=GetParamSoc('SO_APPCO2VALEUR');
          Libelle:=GetParamSoc('SO_APPCO2LIB');
          if not fromContrat then
          begin
          	Part2:=ChargeCodeAppel(i,TypeCode,Valeur,Libelle,FTypeAction,SaisieAppel);
          end else
          begin
          	tmp:=ChargeCodeAppel(i,TypeCode,Valeur,Libelle,FTypeAction,SaisieAppel);
          	if ((TypeCode = 'FIX') and (pos(Tmp,'0; ')=0)) or (TypeCode <>'FIX') then
            begin
              Part2 := Tmp;
            end;
          end;
        End;
      3: Begin
          Code:=Part3;
          TypeCode:=GetParamSoc('SO_APPCO3TYPE');
          Valeur:=GetParamSoc('SO_APPCO3VALEUR');
          Libelle:=GetParamSoc('SO_APPCO3LIB');
          if not fromContrat then
          begin
          	Part3:=ChargeCodeAppel(i, TypeCode,Valeur,Libelle,FTypeAction,SaisieAppel);
          end else
          begin
          	tmp:=ChargeCodeAppel(i,TypeCode,Valeur,Libelle,FTypeAction,SaisieAppel);
          	if ((TypeCode = 'FIX') and (pos(Tmp,'0; ')=0)) or (TypeCode <>'FIX') then
            begin
              Part3 := Tmp;
            end;
          end;
        End;
        else
          Code := '';
        End;
      if ((Typecode='SAI') and (i=2) and (VH_GC.AFFORMATEXER<>'AUC')) then
      BEGIN
        Exercice := Code;
        ExerObligatoire := False;
        if (FTypeAction=taCreat) And (SaisieAppel) then ExerObligatoire := True;
        if Not (InterpreteZoneExercice(VH_GC.AFFORMATEXER,Exercice,AnneeDeb,AnneeFin,Mois,ExerObligatoire)) then
        BEGIN
          PGIBOXAF('Exercice incorrect ',TitreHalley);
          //if Code.CanFocus then Code.SetFocus;
          AviPartErreur:= i;
          Exit;
        END
        else Code:=Exercice;
      END;
    End;
	end;

  // En création d'Appels test qu'une Appel n'existe sur ce code
  // mcd 18/09/01 if Not(ctxScot in V_PGI.PGIContexte) And (FTypeAction=taCreat) And (SaisieAppel) then
  if (GetParamSoc('SO_AFFORMATEXER') ='AUC') And (FTypeAction=taCreat) And (SaisieAppel) then
  BEGIN
    if RechercheAppel (StatutReduit,Part1,Part2,Part3, CodeAvenant,True) then
    BEGIN
      If AfficheMsg Then PGIBoxAF('Code Appel déja utilisé',TitreHalley);
      Exit;
    END;
  END;

  // recup des valeurs des parties du code traité
  Result := RegroupePartiesAppel (StatutReduit, Part1, Part2, Part3, CodeAvenant);

  if (Avenant <> '') then Avenant := CodeAvenant;

End;

//chargement de la partie 1,2 ou 3 du code Appel après repérage découpe.
Function ChargeCodeAppel(Nbpartie : integer;
                         TypeCode, Valeur, Libelle : string;
                         FTypeAction:TActionFiche;
                         SaisieAppel:boolean) : string;
Begin

	  Result := '';

    if ((FTypeAction=taCreat) And (SaisieAppel)) then
        Begin
        	if (Pos(TypeCode,'CPT;CPM')>0) then
          	 Result := CompteurPartieAppel(Nbpartie, false,True)
	        else if (TypeCode ='FIX') then
  	         Result := valeur
    	    else if (TypeCode ='DAT') then
             if length(valeur) = 6 then
	  	    	    Result := FormatDateTime('yyyymm', Now)
             else
	              Result := FormatDateTime('yymm', Now)
        	else
          	 // Si une partie n'est pas renseignée code invalide...
           	Begin
            	 PGIBoxAF('Partie du code Appel non renseignée',TitreHalley);
	             Exit;
           	End;
        End else
        begin
        	result := valeur;
        end;

End;

// En création d'affaire si sous forme de compteur sortie rapide
//car le code sera forcément unique en création
Function RechercheAppel (Aff0,Aff1,Aff2,Aff3,Aff4 : String; EnCreatAff : Boolean):Boolean;
Var stSQL : String;
BEGIN
Result := false;

if (EnCreatAff) and (GetParamSoc('SO_APPCO1TYPE') ='CPT') or
                    (GetParamSoc('SO_APPCO2TYPE') ='CPT') or
                    (GetParamSoc('SO_APPCO3TYPE') ='CPT') then Exit;

if Aff0 = '' then Aff0:='W';
if Aff4 = '' then Aff4:='00';

stSQL :='SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE0="'+ Aff0 + '" AND AFF_AVENANT="'+Aff4+'"';
if (Aff1<>'') then stSQL := stSQL + ' AND AFF_AFFAIRE1="' + Aff1 + '"';
if (Aff2<>'') And (VH_GC.CleAppel.NbPartie > 1) then stSQL := stSQL + ' AND AFF_AFFAIRE2="' + Aff2 + '"';
if (Aff3<>'') And (VH_GC.CleAppel.NbPartie > 2) then stSQL := stSQL + ' AND AFF_AFFAIRE3="' + Aff3 + '"';

Result := ExisteSQL(stSQL);

END;

Function CompteurPartieAppel ( NumPartie : Integer; bProposition,bUpdateCpt : boolean ) : string;
Var NewValeur,Valeur,suffixe : String;
    Cpt,lng : integer;
    Q : TQuery;
Begin

  Result := ''; NewValeur := ''; Valeur := '';
  Suffixe := '';

  case NumPartie of
      1: lng := GetParamSoc('SO_APPCO1LNG');
      2: lng := GetParamSoc('SO_APPCO2LNG');
      3: lng := GetParamSoc('SO_APPCO3LNG');
      else
      lng := 0;
  End;

  Q := nil;
  try
    Q := OpenSQL('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_APPCO' + IntToStr(NumPartie) + 'VALEUR'+ suffixe + '"',True,-1,'',true);
    if Not Q.EOF then Valeur:= Q.Fields[0].asstring;

    // sortie rapide si une erreur est rencontrée
    if Valeur = '' then begin Result := ''; exit end;

    if length(Valeur) < lng then
        BEGIN
        Cpt := StrToInt(Valeur);
        Valeur := Format('%*.*d',[lng,lng,Cpt]);
        END;
    Result:= Valeur;

    NewValeur := CodeAppelSuivant(Valeur, lng);

    //mcd 03/08/00 cas ou compteur faux ou depasse X éléments en démo
    if (NewValeur = '') then begin
       result:='';
       exit;
       end;

    // MAJ du compteur incrémenté en mémoire + Table
    if bUpdateCpt then
       begin
       SetParamsoc('SO_APPCO'+IntToStr(NumPartie)+'VALEUR'+Suffixe, NewValeur );
       ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA="'+NewValeur+'" WHERE SOC_NOM="SO_APPCO' + IntToStr(NumPartie) + 'VALEUR'+ suffixe + '"');
       end
    else
       begin
    {$IFDEF DEBUG}
       PGIBoxAF('Attention valeur cpt affaire non enregistré n°: ' + NewValeur,'Debug Compteur affaire');
    {$ENDIF}
       end;


       case NumPartie of
            1:  SetParamSoc('SO_APPCO1VALEUR', NewValeur) ;
            2:  SetParamSoc('SO_APPCO2VALEUR', NewValeur) ;
            3:  SetParamSoc('SO_APPCO3VALEUR', NewValeur) ;
       End;

  finally
    Ferme(Q);
  end;
End;

function InterpreteZoneExercice (TypeExer : string;var Exercice : String; Var  AnneeDeb,AnneeFin,Mois : integer; Obligatoire : Boolean) : Boolean;
var
    tmp : string;
BEGIN
     // MCD 13/11/00 MOis = Mois de fin (et non pas début)
Result := True;
   // mcd 18/09/01 if Not (ctxScot in V_PGI.PGIContexte) then Exit;
if TypeExer ='AUC' then Exit;
if Not(Obligatoire) and (Exercice = '') then Exit;
AnneeDeb := 0;  AnneeFin := 0; Mois := 0;
  // mcd 18/09/01 if TypeExer = 'AUC' then Exit ;
if Exercice = '' then BEGIN Result := False; Exit; END;

// Test si le nombre de caractères à été bien renseigné
Exercice := Trim (Exercice);
if (TypeExer = 'AM')And (Length(Exercice)<> 6) then BEGIN Result := False; Exit; END;
if (TypeExer = 'AA')And (Length(Exercice)<> 8)And (Length(Exercice)<> 4) then BEGIN Result := False; Exit; END;
if (TypeExer = 'A') And (Length(Exercice)<> 4) then BEGIN Result := False; Exit; END;

// interprétation des zones saisies
if TypeExer = 'AM'  then
    BEGIN
    // mcd 13/11/00 tmp := Copy(Exercice,1,4); if trim(tmp) <> '' then AnneeDeb := StrToInt(tmp ) else AnneeDeb:= 0;
    tmp := Copy(Exercice,1,4); if trim(tmp) <> '' then AnneeFin := StrToInt(tmp ) else AnneeFin:= 0;
    tmp := Copy(Exercice,5,2); if trim(tmp) <> '' then Mois := StrToInt(tmp) else mois := 0;
    END else
if TypeExer = 'AA'  then
    BEGIN
    tmp := Copy(Exercice,1,4); if trim(tmp) <> '' then AnneeDeb := StrToInt(tmp ) else AnneeDeb:= 0;
    tmp := Copy(Exercice,5,4);
    if trim(tmp)<> '' then AnneeFin := StrToInt(tmp)
                      else BEGIN AnneeFin:= AnneeDeb; Exercice :=Copy(Exercice,1,4)+Copy(Exercice,1,4); END;
                        // Année de fin d'exercice = Année de début d'exercice par défaut
    END else
if TypeExer = 'A'  then
    BEGIN
    tmp := Copy(Exercice,1,4); if trim(tmp) <> '' then AnneeFin := StrToInt(tmp ) else AnneeFin:= 0;
    END;

// En attente ...
(*if (TypeExer = 'XX') or (TypeExer = 'XX') then  // Test de l'année début sur 2 car
    BEGIN
    if (AnneeDeb < 0) Or (AnneeDeb > 99) then BEGIN Result := False; AnneeDeb := 0; END;
    if (AnneeDeb > 90) then  AnneeDeb := AnneeDeb + 1900 else AnneeDeb := AnneeDeb + 2000;
    END; **)
// Test de cohérence
if (TypeExer = 'AM') then  // Test du Mois
    BEGIN
    if (Mois < 1) Or (Mois > 12) then BEGIN Result := False; Mois := 0; END;
    if (Mois <> 1) then AnneeDeb := AnneeFin - 1 else AnneeDeb := AnneeFin;
    END;
if (TypeExer = 'AA') then  // Test comparatif des 2 années
    BEGIN
    if (AnneeDeb > AnneeFin) then BEGIN AnneeFin := AnneeDeb; Result := false; END;
    END;

END;

Function RegroupePartiesAppel (Var Part0, Part1, Part2, Part3,Avenant : String) : string;
Var Valeur : string;
    Lng : integer;
    NbPartie : integer;
BEGIN
Valeur := '';
lng := 0;

Avenant := BTPTestCodeAvenant(Avenant,True);
Part0 := BTPTestStatutReduit(Part0); // 1er car. correspondant au statut
Valeur := Part0;
NbPartie := GetParamSoc('SO_APPCODENBPARTIE');

If (NbPartie >= 1) then
   Begin
   lng := GetParamSoc('SO_APPCO1LNG');
   Valeur:=Valeur + Format('%-*.*s',[lng,lng,Part1]);
   end;

If (NbPartie >= 2) then
   Begin
   lng := GetParamSoc('SO_APPCO2LNG');
   Valeur:=Valeur + Format('%-*.*s',[lng,lng,Part2]);
   end;

If (NbPartie >= 3) then
   Begin
   lng := GetParamSoc('SO_APPCO3LNG');
   Valeur:=Valeur + Format('%-*.*s',[lng,lng,Part3]);
   end;

Lng := 15 - Length(Valeur);
if Lng > 0 then Valeur := Valeur + Format('%-*.*s',[lng,lng,'']);
valeur := valeur + Avenant;

Result:=TrimRight(Valeur);

END;

function CodeAppelSuivant ( Code : String; Lng : Integer ) : String ;
Var i ,cpt: integer ;
    Qr : TQuery;
    nbr:integer;
BEGIN

if code = '' then code := '0';

Code:=AnsiUppercase(Trim(Code)) ;

if Copy(code,1,lng) < Copy('99999999999999999',1,lng) then
    BEGIN
    // code numerique, on incrémente
    Cpt := StrToInt(code);
    Inc(Cpt);
    //mcd 03/08/00
    if V_PGI.VersionDemo then
       begin
       Qr:=OpenSQL('SELECT COUNT(*) FROM AFFAIRE WHERE AFF_AFFAIRE0="W"',True,-1,'',true) ;
       if Not Qr.EOF then
          Nbr:=Qr.Fields[0].AsInteger
       else
          Nbr:=0;
       Ferme(Qr) ;
       if (Nbr > 50) then
          begin
          PGIBoxAf('Version démonstration,Seulement 50 fiches appels permises',TitreHalley);
          Result:='' ;
          exit;
          end;
       end;
    Code := Format('%*.*d',[lng,lng,Cpt]);
    Result:=Code ;
    END
else
    BEGIN
    // numérotation alpha des compteurs
    if Copy(code,1,lng) < Copy('ZZZZZZZZZZZZZZZZZ',1,lng) then
       begin
       if Length(Code) < Lng then
          Result:=Copy('AAAAAAAAAAAAAAAAA',1,Lng)
       else
          BEGIN
          i:=Lng ;
          While Code[i]='Z' do
             BEGIN
             Code[i]:='A' ;
             Dec(i) ;
             END;
         Code[i]:=Succ(Code[i]) ;
         // ; pose pb ds traitements !!!
         if Code[i]= ';' then Code[i]:=Succ(Code[i]) ;
         Result:=Code ;
         END;
       end
    else
       begin
          // Si arrivé au maxi, on considère que l'on peut repartir à 0
          // et que les missions ayant ces valeurs de compteurs seront supprimées
          // attention si compteur sur 3c ...
         Result:=Copy('00000000000000',1,lng);
       end;
    END;

END;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 02/02/2000
Modifié le ... : 17/02/2000
Description .. : Décomposition du code appel concaténé
Mots clefs ... : AFFAIRE;CODE AFFAIRE
*****************************************************************}
Procedure CodeAppelDecoupe(CodeApp : string; Var Part0, Part1, Part2, Part3, Avenant : String);
var i,lng,posit: integer;
    TexteSel,CodePartie : String;
Begin

  // mcd 28/03/03 passage de tous les VH_GC en GetPAramSoc
  Part1 := ''; Part2 := ''; Part3 := '';
  Posit := 0;  lng := 0;

  If (CodeApp = '') then Exit;

  // mcd 28/03/03 for i:=1 to NbMaxPartieAffaire do
  for i:=1 to GetParamSoc('SO_APPCodeNbPartie') do
      Begin
      // recup des valeurs de la partie du code traitée
      case i of
           1: Begin
              lng:=GetParamSoc('SO_APPCo1Lng');
              posit:=2;
              End;
        // Attention le premier car. réservé pour le statut
        	 2: Begin
           		lng:= GetParamSoc('SO_APPCo2Lng');
           		posit:=GetParamSoc('SO_APPCo1Lng')+2;
           		End;
        	 3: Begin
           	  lng:=GetParamSoc('SO_APPCo3Lng');
              posit:=GetParamSoc('SO_APPCO1LNG')+GetParamSoc('SO_APPCO2LNG')+2;
              End;
        End;

    // Attention  longueur Maxi : 15 (car 16 et 17 = avenant )
    if Posit + Lng > 16 then Lng := 16 -Posit;

    If (i <= GetParamSoc('SO_APPCodeNbPartie')) Then
        Begin   // Partie utilisée
        TexteSel:='';
        If(CodeApp<> '') Then TexteSel:= trim(Copy (CodeApp,posit,lng));
        CodePartie:= TexteSel;
        End
    Else
        Begin  // RAZ de la partie de structure non utilisée
        CodePartie:='';
        End;
    case i of
        1: Part1 := CodePartie;
        2: Part2 := CodePartie;
        3: Part3 := CodePartie;
        End;
    End;

  // Traitement du statut et de la zone avenant
  Part0 := BTPTestStatutreduit(Copy(CodeApp,1,1));
  Avenant := BTPTestCodeAvenant(Copy (codeApp,16,2),False);

End;


{$IFDEF EAGLCLIENT}
Procedure TraiteNomAppelGrid ( Gr : THGrid; Nom : string; Col : integer);
Var  Lg,ind,NumPart : integer;
     NomChamp,Libelle : string;
     Visible : Boolean;
begin
if nom = '' then Exit;
lg := Length(Nom); ind:= Pos ('_',Nom);
if (lg =0) or (ind = 0) then  exit;
Nomchamp := Copy (Nom, ind+1,lg-ind);
Visible:=False; //mcd 14/02/03
If ((Nomchamp = 'AFFAIRE1') Or (Nomchamp = 'AFFAIRE2') Or (Nomchamp = 'AFFAIRE3')) Then
  Begin
  NumPart := StrToInt(Copy(NomChamp, 8,1));
  Case NumPart Of
      1 : Begin Visible:=True; Libelle:=VH_GC.CleAppel.Co1Lib;  End;
      2 : Begin Visible:=VH_GC.CleAppel.Co2Visible; Libelle:=VH_GC.CleAppel.Co2Lib; End;
      3 : Begin Visible:=VH_GC.CleAppel.Co3Visible; Libelle:=VH_GC.CleAppel.Co3Lib; End;
      End;
  if not(Visible) then Gr.colwidths[Col]:=0;
  Gr.Cells[Col,0]:=Libelle;
  End;
if ((Nomchamp='AVENANT') and Not(VH_GC.CleAppel.GestionAvenant)) then Gr.colwidths[Col]:=0;
end;

Procedure ModifColAppelGrid (Gr : THGrid; TobQ : TOB );
Var  i,PosFr   : integer;
     Nom,stSQL : string;
begin
if (TOBQ = Nil) then Exit;
i := 0;
if TobQ.FieldCount = 0 then
   begin   // Aucun enregistrement trouvé ...
   stSQL := TOBQ.SQL;
   if stSQL = '' then exit;
   PosFr := Pos ('FROM',stSQL); if PosFr = 0 then Exit;
   stSQL := Copy (stSQL,8,PosFr-8);
   Nom:=(Trim(ReadTokenPipe(stSQL,',')));
   While (Nom <>'') do
      begin
      TraiteNomAppelGrid (Gr,Nom,i);
      Nom:=(Trim(ReadTokenPipe(stSQL,',')));
      Inc(i);
      end;
   end
else
   begin
   for i:=0 to TobQ.FieldCount-1 do
      BEGIN
      Nom := TOBQ.Fields[i].FieldName;
      TraiteNomAppelGrid (Gr,Nom,i);
      END;
   end;
end;

{$ELSE}
Procedure ModifColAppelGrid (Gr : THDBGrid );
var i, NumPart,ind,lg : integer;
    visible : Boolean;
    Libelle,Nomchamp : String;
Begin
If Not(Gr is THDBGrid) Then Exit;
if (Gr.Columns.Count = 1) then Exit ;
Visible :=False;
For i:=0 to Gr.Columns.Count-1 do
    Begin
    lg := Length(Gr.Columns[i].FieldName);
    ind := Pos ('_',Gr.Columns[i].FieldName);
   // mcd 30/11/01 car sort si libellé au lieu code !!! if ((ind = 0) or (lg =0)) then  exit;
    if ((ind = 0) or (lg =0)) then  continue;
    Nomchamp := Copy (Gr.Columns[i].FieldName,ind+1, lg-ind);

    If ((Nomchamp = 'AFFAIRE1') Or (Nomchamp = 'AFFAIRE2') Or (Nomchamp = 'AFFAIRE3')) Then
        Begin
        NumPart := StrToInt(Copy(NomChamp, 8,1));
        Case NumPart Of
            1 : Begin
            			Visible:=True;
                  Libelle:=VH_GC.CleAppel.Co1Lib;
            		End;
            2 : Begin Visible:=VH_GC.CleAppel.Co2Visible; Libelle:=VH_GC.CleAppel.Co2Lib; End;
            3 : Begin Visible:=VH_GC.CleAppel.Co3Visible; Libelle:=VH_GC.CleAppel.Co3Lib; End;
            End;
        Gr.columns[i].visible:=visible ;
              // mcd 25/07/02 mis en majuscule pour éviter pb de frappe min/maj
        If ((AnsiUppercase(Copy(Gr.Columns[i].Field.DisplayLabel, 1, 6)) = 'PARTIE'))or
          ((Copy(Gr.Columns[i].Field.DisplayLabel, 1,17) = 'Code appel partie'))or            // mcd 30/05/02 dans GL, les libelles ne sont pas les mêmes que dan sles autres tables !!!!!
          ((Copy(Gr.Columns[i].Field.DisplayLabel, 1,17) = 'Code Appel partie'))Then            // mcd 30/05/02 dans Gp , les libelles ne sont pas les mêmes que dan sles autres tables !!!!!
            Gr.columns[i].Field.DisplayLabel:=Libelle ;   // valeur par défaut non modifiée par l'utilisateur
            // Attention dans les listes bien appeler les découpages du code affaire partie 1 , 2, 3
        End;
{$IFNDEF BTP}
    // Gestion de la colonne Avenant
    if (Nomchamp = 'AVENANT') then
        BEGIN
        if Not (VH_GC.CleAppel.GestionAvenant) then  Gr.columns[i].visible:= False;
        END;
{$ENDIF}
    End;
FormatExerciceGrid(Gr);
End;

function FormatNoTel(Notel : String) : string;
Var Car       : String;
		i					: integer;
Begin

  for i:=1 to length(NoTel) do
	  begin
    Car:=copy(NoTel,i,1);
    if (IsNumeric(car)) and (car <>'.') and (car<>',') and (car<>'-') and (car<>' ') then
       Result:=Result+car;
  	end;

	if GetParamSoc('SO_RTSUPPZERO') then Result := '0'+ Result;

end;

//Calcul du numéro d'appel en création
Function CalculCodeAppel(P0,P1,P2,P3,Avn, Tiers : String; FromContrat : boolean=false) : String;
var tmp           : string;
    iPartErreur   : integer;
begin
  result := '';
	tmp := DechargeCleAppel (P0,P1,P2,P3,Avn,Tiers, TaCreat, True, True, false, iPartErreur,FromContrat);
  if tmp <> '' then Result := tmp;
end;

{$ENDIF}


end.
