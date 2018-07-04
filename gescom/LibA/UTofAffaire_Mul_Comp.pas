unit UTofAffaire_Mul_Comp;

interface
{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : AFFAIRE_MUL_COMPTA : Tof associée au mul de sélection affaire pour la compta.
Suite ........ : repris de la tof de la GI, pour l'epurer et mettre pleins de
Suite ........ : choses en dur, afin que la compta ne link pas avec toute la
Suite ........ : GI
Mots clefs ... : AFFAIRE;COMPTA
*****************************************************************}

uses  StdCtrls,Controls,Classes,M3FP,
{$IFDEF EAGLCLIENT}
      eMul,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,DBGrids,HDB,
{$ENDIF}
      forms,sysutils, ParamSoc,
      ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF, UTOB,lookup,UtilPgi;

Type
     TOF_Affaire_Mul_Comp = Class (TOF)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnUpdate ; override ;
        procedure OnLoad ; override ;
        private
          bChangeTiers   :  Boolean;
          Procedure ChargeCleAffaireCom(Part0,Part1,Part2,Part3,Avenant: THCritMaskEDIT;{ FTypeAction:TActionFiche; }CodeAff:string);
          procedure OnElipsisClickAffairePart1 (Sender : TObject );
          procedure OnElipsisClickAffairePart2 (Sender : TObject );
          procedure OnElipsisClickAffairePart3 (Sender : TObject );
     END ;

function FormatZoneExerciceCom (TypeExer : string; bTiretStocke : Boolean = False) : String;
{$IFDEF EAGLCLIENT}
Procedure TraiteNomAffaireGridComp ( Gr : THGrid; Nom : string; Col : integer);
Procedure ModifColAffaireGridComp (Gr : THGrid; TobQ : TOB );
{$ELSE}
Procedure ModifColAffaireGridComp (Gr : THDBGrid );
Procedure FormatExerciceGridCOmp ( Gr : THDBGrid ) ;
{$endif}
function GetParamSocGI ( Nom : string ) : variant;
function ChercheLibelleGCZoneLibre ( Code : string ) : string;
implementation


{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Onargument : peut recevoir STATUT, mais pour l'instant non géré et figé à AFF
Suite ........ : Peut recevoir l'auxiliaire
Suite ........ : ne traite pas les sous affaire
Suite ........ : GI
Mots clefs ... : AFFAIRE;COMPTA
*****************************************************************}
procedure TOF_Affaire_Mul_Comp.OnArgument(stArgument : String );
var Critere,Statut,Staff : string;
Part0,Part1,Part2,Part3,Avenant: ThEdit;
Q : TQuery;
i:integer;
TableTiers : string;
Begin
Inherited;
Statut :='';
PArt0 :=THEdit(GetControl('AFF_AFFAIRE0'));
PArt1 :=THEdit(GetControl('AFF_AFFAIRE1'));
PArt2 :=THEdit(GetControl('AFF_AFFAIRE2'));
PArt3 :=THEdit(GetControl('AFF_AFFAIRE3'));
Avenant :=THEdit(GetControl('AFF_AVENANT'));
bChangeTiers := True;
Statut := 'AFF';
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
   if (copy(Critere,1,4) ='AUXI')
	then begin
             // recu code auxiliaire. On cherche le tiers correspondant
        if GetParamSoc('SO_CPSAISIEAFFAIRE')='COM' then TableTiers := '##DP##.TIERS'
        else TableTiers := 'TIERS';
        Staff:='';
        staff:='SELECT T_TIERS From '+TableTiers+' WHERE T_AUXILIAIRE="'+Copy(Critere,6,Length(critere)-5)+'"';
        Q:=OPENSQL('SELECT T_TIERS From '+TableTiers+' WHERE T_AUXILIAIRE="'+Copy(Critere,6,Length(critere)-5)+'"',True);
        If Not Q.EOF Then Staff:=Q.Fields[0].AsString;
        Ferme(Q);
        SetControlEnabled ('aff_tiers',False);
        SetCOntrolText('AFF_TIERS',Staff);
        SetControlVisible ('BEffaceaff1',False);
	end;
    if Copy(Critere,1,6) = 'STATUT' then
      begin    // en principe non gérer, uniquement AFF
      if copy(critere,8,3)= 'PRO' then begin
         SetControlProperty('AFF_ETATAFFAIRE','Plus',' AND (CC_LIBRE="" OR CC_LIBRE="PRO")');
         Statut := 'PRO';
         end
      else begin
         SetControlProperty('AFF_ETATAFFAIRE','Plus',' AND (CC_LIBRE="" OR CC_LIBRE="AFF")');
         Statut := 'AFF';
         end;
      end;
    if critere = 'NOCHANGESTATUT' then
        BEGIN
        SetControlVisible ('AFF_STATUTAFFAIRE',False);
        SetControlVisible ('TAFF_STATUTAFFAIRE',False);
        END;
    if critere = 'NOCHANGETIERS' then BEGIN SetControlEnabled('AFF_TIERS',False); bChangetiers := False; END;
    if critere = 'MODELEONLY' then SetControlProperty ('AFF_MODELE','Checked',True);
    if critere = 'NOFILTRE' then   Tfmul(Ecran).FiltreDisabled:=true;
    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;
SetControlVisible ('BInsert',False);
if not(GetParamSocGI('SO_AffCo2Visible')) then SetCOntrolVisible('AFf_AFFAIRE2',False);
if not(GetParamSocGI('SO_AffCo3Visible')) then SetCOntrolVisible('AFf_AFFAIRE3',False);
if not(GetParamSocGI('SO_AffGestionAvenant')) then SetCOntrolVisible('AFf_AVENANT',False);
// Caption de la form
if Statut = 'PRO' then Ecran.Caption := 'Propositions d''affaires';
UpdateCaption(Ecran);
ChargeCleAffaireCom (Part0,Part1,Part2,Part3,Avenant,{afActionFiche,}StAff);
 // on traduit les libelle champ libre
For I :=1 to 9 do begin
//    Staff:= RechDom('GCZONELIBRE','MT'+InttoStr(i),False);
      Staff:= ChercheLibelleGCZoneLibre('MT'+InttoStr(i));
    if (Staff <>'.') and (STaff <> '-') then SetCOntrolText('TAFF_LIBREAFF'+InttoStr(i),Staff )
       else begin
            SetCOntrolVisible('TAFF_LIBREAFF'+InttoStr(i),False);
            SetCOntrolVisible('AFF_LIBREAFF'+InttoStr(i),False);
            end;

    end;
//Staff:= RechDom('GCZONELIBRE','MTA',False);
Staff:= ChercheLibelleGCZoneLibre('MTA');
if (Staff <>'.') and (STaff <> '-') then SetCOntrolText('TAFF_LIBREAFFA',Staff )
   else begin
        SetCOntrolVisible('TAFF_LIBREAFFA',False);
        SetCOntrolVisible('AFF_LIBREAFFA',False);
        end;
For I :=1 to 3 do begin
//    Staff:= RechDom('GCZONELIBRE','MR'+InttoStr(i),False);
    Staff:= ChercheLibelleGCZoneLibre('MR'+InttoStr(i));
    if (Staff <>'.') and (STaff <> '-') then SetCOntrolText('TAFF_RESSOURCE'+InttoStr(i),Staff )
       else begin
            SetCOntrolVisible('TAFF_RESSOURCE'+InttoStr(i),False);
            SetCOntrolVisible('AFF_RESSOURCE'+InttoStr(i),False);
            end;
    end;
End;

{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 17/09/2001
Modifié le ... :   /  /
Description .. : Fct repris de chargeCleAffaire, mais mis dans le spource, 
Suite ........ : pour ne pas obliger la compta à linker avec toute la GI/GA
Mots clefs ... : AFFAIRE;CLE;COMPTA
*****************************************************************}
Procedure TOF_Affaire_Mul_Comp.ChargeCleAffaireCom(Part0,Part1,Part2,Part3,Avenant: THCritMaskEDIT;{ FTypeAction:TActionFiche;} CodeAff:string);
var i,lng, PositAvenant,NBMaxPartieAffaire : integer;
    TypeCode : string3;
    Valeur,Libelle,CodePartie0,CodePartie1,CodePartie2,CodePartie3,CodeAvenant : String;
    Visible : Boolean;
    Code : THCritMaskEdit;
    OnChangetmp1, OnChangetmp2 : TNotifyEvent;
Begin

OnChangetmp1 := Part1.OnChange;
OnChangetmp2 := Part2.OnChange;
Part1.OnChange := Nil;
Part2.OnChange := Nil;
NbMaxPartieAffaire:=3;   // nbre maxi des parties codes affaire : 3
try
PositAvenant := 0;
// Gestion des affaires
if Part0 <> Nil then
    BEGIN
    Part0.DataType:= 'AFSTATUTREDUIT';
    Part0.ElipsisButton:=True;
    END;

for i:=1 to NbMaxPartieAffaire do
    Begin
    // recup des valeurs de la partie du code traitée
    case i of
        1: Begin Code:=Part1;
                 lng:=StrToInt(GetParamSocGI('SO_AFFCo1Lng'));
                 TypeCode:=GetParamSocGI('SO_AFFCo1Type');
                 Valeur:=GetParamSocGI('SO_AFFCo1valeur');
                 Visible:=True;
                 Libelle:=GetParamSocGI('SO_AFFCo1Lib');
                 End;
        2: Begin Code:=Part2;
                 lng:=StrToInt(GetParamSocGI('SO_AFFCo2Lng'));
                 TypeCode:=GetParamSocGI('SO_AFFCo2Type');
                 Valeur:=GetParamSocGI('SO_AFFCo2valeur');
                 Visible:=GetParamSocGI('SO_AFfCo2Visible');
                 Libelle:=GetParamSocGI('SO_AFFCo2Lib');
                 End;
        3: Begin Code:=Part3;
                 lng:=StrToInt(GetParamSocGI('SO_AFFCo3Lng'));
                 TypeCode:=GetParamSocGI('SO_AFFCo3Type');
                 Valeur:=GetParamSocGI('SO_AFFCo3valeur');
                 Visible:=GetParamSocGI('SO_AFfCo3Visible');
                 Libelle:=GetParamSocGI('SO_AFFCo3Lib');
                 End;
        End;
    If (i <= StrToInt(GetParamSocGI('SO_AFFCodeNbPartie'))) Then
        Begin   // Partie utilisée
        Code.Visible:= Visible; Code.Hint:=Libelle; // propriété visible + conseil
        Code.ElipsisButton:=False;
        Code.MaxLength:=lng;  Code.Width :=(lng * 8)+10; // Calcul longueur des zones
        if ((Typecode='SAI') And (i=2) And (GetParamSocGI('SO_AFFORMATEXER')<>'AUC')) then
           Code.Width := Code.Width + 15; // séparateurs en plus
        If (i=2) Then Code.Left:=Part1.Left+Part1.Width+5;
        If (i=3) Then Code.Left:=Part2.Left+Part2.Width+5;
        Code.ReadOnly:=False;
        // spécificitées en fonction du type de zone ...
        if (Typecode= 'LIS') then                                       // Liste de saisie
            Begin
            // Code.DataType:='AFFAIREPART'+intToStr(i);
            Code.ElipsisButton:=True;
            case i of
              1 : Code.OnElipsisClick := OnElipsisClickAffairePart1;
              2 : Code.OnElipsisClick := OnElipsisClickAffairePart2;
              3 : Code.OnElipsisClick := OnElipsisClickAffairePart2;
            end;
            Code.Width:=45;
            End else
        if ((Typecode='SAI') And (i=2) And (GetParamSocGI('SO_AFFormatExer')<>'AUC')) then  // Formattage Exercice
            BEGIN
            Code.EditMask := FormatZoneExerciceCom(GetParamSocGI('SO_AFFormatExer'),false);
            if (GetParamSocGI('SO_AFFormatExer')='AA') then
                Code.Hint :=Code.Hint+ '(Année Début-Année fin d''exercice)'
            else if (GetParamSocGI('SO_AFFormatExer')='A') then
                Code.Hint :=Code.Hint+ '(Millésime : Année de fin d''exercice)'
            else
                Code.Hint :=Code.Hint+ '(Année Début-Mois Début d''exercice)';
            END;
            // mcd déplacer, car longueur change si LIS 26/06/01
        if Visible then
           BEGIN
           PositAvenant := Code.Left +  Code.Width + 5; // Calcul position avenant
           END;
        End
    Else
        Begin  // RAZ de la partie de structure non utilisée
        Code.Visible:=False;
        End;
    End;
CodePartie1 := ''; CodePartie2 := ''; CodePartie3 :=''; CodePartie0 := ''; CodeAvenant := '';
{if (codeAff <> '')  then
    CodeAffaireDecoupe(CodeAff,CodePartie0,CodePartie1,CodePartie2,CodePartie3,CodeAvenant,FTypeAction,SaisieAffaire);
  }
if (Part0<>nil) then
   Part0.Text :=CodePartie0 ;
Part1.Text :=CodePartie1 ; Part2.Text := CodePartie2; Part3.Text := CodePartie3;

// Gestion de la zone avenant
if Avenant <> Nil then
    BEGIN
    if (GetParamSocGI('SO_AFFGestionAvenant')) then
        BEGIN
        Avenant.Visible := true;
        Avenant.Left := PositAvenant; Avenant.Width :=2 * 10 + 5;
        Avenant.Hint := 'Avenant';    Avenant.MaxLength := 2;
      //  Avenant.Text := TestcodeAvenant(CodeAvenant,SaisieAffaire);
        END
    else Avenant.Visible := false;
    END;
finally
Part1.OnChange:= OnChangetmp1;
Part2.OnChange:= OnChangetmp2;
end;
End;

{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 17/09/2001
Modifié le ... :   /  /    
Description .. : Fct dupliquer pour la compta depuis Fct de affair util, afin de
Suite ........ : ne pas avoir à linker la compta avec toute la GI
Mots clefs ... : EXER;AFFAIRE
*****************************************************************}
function FormatZoneExerciceCom (TypeExer : string; bTiretStocke : Boolean = False) : String;
var tiret : string;
BEGIN
Result := '';
if bTiretStocke then tiret :='1' else tiret :='0';
if TypeExer = 'AUC' then Exit else
if TypeExer = 'A' then Exit else
if TypeExer = 'AM'  then Result := '!9999\-00;'+tiret+';_' else
if TypeExer = 'AA' then Result := '!9999\-0000;'+tiret+';_'
;
END;

procedure TOF_Affaire_Mul_Comp.OnUpdate;
Begin
// Gestion repositionnement auto sur l'affaire en cours si sortie rapide / bug par prg ( ou Eagl)

inherited;
if not(GetParamSocGI('SO_AffCo2Visible')) then SetControlText('AFF_AFFAIRE2','');
if not(GetParamSocGI('So_AffCo3Visible')) then SetControlText('AFF_AFFAIRE3','');
{$IFDEF EAGLCLIENT}
if (Ecran is TFMul) then ModifColAffaireGridComp ( TFMul(Ecran).FListe,TFMul(Ecran).Q.tq);
{$ELSE}
if (Ecran is TFMul) then ModifColAffaireGridComp ( TFMul(Ecran).FListe);
{$ENDIF}
End;



procedure TOF_Affaire_Mul_Comp.OnLoad;
var F : TFMul ;
begin
  if not (Ecran is TFMul) then exit ;
  F:=TFMul(Ecran) ;
  F.Q.Manuel:=True ; // Evite l'exécution de la requête lors de la maj de Q.Liste
  // Affectation de la liste ad-hoc
  if GetParamSoc('SO_CPSAISIEAFFAIRE')='COM' then
    F.Q.Liste := 'AFMULRECHAFFCPDP'
  else F.Q.Liste := 'AFMULRECHAFFCOMPT';
  F.Q.Manuel:=False ;
  // suppression de la gestion de sous affaire fait en GI/GA
end;

{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 17/09/2001
Modifié le ... : 17/09/2001
Description .. : Modifcation dynamique des colonnes d'un grid en fonction du paramétrage du code affaire por la compta
Mots clefs ... : AFFAIRE;CODE AFFAIRE
*****************************************************************}
{$IFDEF EAGLCLIENT}
Procedure TraiteNomAffaireGridComp ( Gr : THGrid; Nom : string; Col : integer);
Var  Lg,ind,NumPart : integer;
     NomChamp,Libelle : string;
     Visible : Boolean;
begin
if nom = '' then Exit;
lg := Length(Nom); ind:= Pos ('_',Nom);
if (lg =0) or (ind = 0) then  exit;
Nomchamp := Copy (Nom, ind+1,lg-ind);

If ((Nomchamp = 'AFFAIRE1') Or (Nomchamp = 'AFFAIRE2') Or (Nomchamp = 'AFFAIRE3')) Then
  Begin
  NumPart := StrToInt(Copy(NomChamp, 8,1));
  Case NumPart Of
      1 : Begin Visible:=True; Libelle:=GetParamSocGI('SO_AffCo1Lib');  End;
      2 : Begin Visible:=GetParamSocGI('SO_AFFCo2Visible'); Libelle:=GetParamSocGI('SO_AFFCo2Lib'); End;
      3 : Begin Visible:=GetParamSocGI('SO_AFFCo3Visible'); Libelle:=GetParamSocGI('SO_AFFCo3Lib'); End;
      End;
  if not(Visible) then Gr.colwidths[Col]:=0;
  Gr.Cells[Col,0]:=Libelle;
  End;
// Gestion de la colonne Avenant
if ((Nomchamp='AVENANT') and Not(GetParamSocGI('SO_AFFGestionAvenant'))) then Gr.colwidths[Col]:=0;
end;

Procedure ModifColAffaireGridComp (Gr : THGrid; TobQ : TOB );
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
      TraiteNomAffaireGridComp (Gr,Nom,i);
      Nom:=(Trim(ReadTokenPipe(stSQL,',')));
      Inc(i); 
      end;
   end
else
   begin
   for i:=0 to TobQ.FieldCount-1 do
      BEGIN
      Nom := TOBQ.Fields[i].FieldName;
      TraiteNomAffaireGridComp (Gr,Nom,i);
      END;
   end;
end;

{$ELSE}
Procedure ModifColAffaireGridComp (Gr : THDBGrid );
var i, NumPart,ind,lg : integer;
    visible : Boolean;
    Libelle,Nomchamp : String;
Begin
If Not(Gr is THDBGrid) Then Exit;
if (Gr.Columns.Count = 1) then Exit ;

For i:=0 to Gr.Columns.Count-1 do
    Begin
    lg := Length(Gr.Columns[i].FieldName);
    ind := Pos ('_',Gr.Columns[i].FieldName);
    if ((ind = 0) or (lg =0)) then  exit;
    Nomchamp := Copy (Gr.Columns[i].FieldName,ind+1, lg-ind);

    If ((Nomchamp = 'AFFAIRE1') Or (Nomchamp = 'AFFAIRE2') Or (Nomchamp = 'AFFAIRE3')) Then
        Begin
        NumPart := StrToInt(Copy(NomChamp, 8,1));
        Case NumPart Of
            1 : Begin Visible:=True; Libelle:=GetParamSocGI('SO_affCo1Lib');  End;
            2 : Begin Visible:=GetParamSocGI('SO_AFFCo2Visible'); Libelle:=GetParamSocGI('SO_AFFCo2Lib'); End;
            3 : Begin Visible:=GetParamSocGI('SO_AFFCo3Visible'); Libelle:=GetParamSocGI('SO_AFFCo3Lib'); End;
            End;
        Gr.columns[i].visible:=visible ;
        If (Copy(Gr.Columns[i].Field.DisplayLabel, 1, 6) = 'Partie') Then
            Gr.columns[i].Field.DisplayLabel:=Libelle ;   // valeur par défaut non modifiée par l'utilisateur
            // Attention dans les listes bien appeler les découpages du code affaire partie 1 , 2, 3
        End;
    // Gestion de la colonne Avenant
    if (Nomchamp = 'AVENANT') then
        BEGIN
        if Not (GetParamSocGI('SO_AFFGestionAvenant')) then  Gr.columns[i].visible:= False;
        END;
    End;
FormatExerciceGridComp(Gr);
End;
Procedure FormatExerciceGridCOmp ( Gr : THDBGrid ) ;
Var i : integer;
BEGIN
if Not(ctxScot in V_PGI.PGIContexte) then exit;
For i:=0 to Gr.Columns.Count-1 do
    BEGIN
    if (Gr.Columns[i].FieldName = 'AFF_AFFAIRE2') then
        Gr.columns[i].Field.EditMask := FormatZoneExerciceCom(GetParamSocGI('SO_AFFORMATEXER'),false);
    END;
END;
{$ENDIF}

function GetParamSocGI ( Nom : string ) : variant;
var Q : TQuery;
  St, Data,TableSoc : string;
  zz : Char;
begin
  if GetParamSoc('SO_CPSAISIEAFFAIRE')='COM' then
  begin
       // lit tout pour un seul enrgt paramsoc. on laisse
    Q:=OpenSQL('SELECT * FROM ##DP##.PARAMSOC WHERE SOC_NOM="'+AnsiUppercase(Trim(Nom))+'"',TRUE) ;
    if Not Q.EOF then
    BEGIN
      St:=Q.FindField('SOC_DESIGN').AsString ;
      St:=ReadTokenSt(St) ; if St<>'' then zz:=St[1] else zz:=' ' ;
      Data:=Q.FindField('SOC_DATA').AsString ;
      Result:=ParamSocDataToVariant(Nom,Data,zz) ;
    END else
    BEGIN
      result:=#0 ;
{$IFNDEF EAGLSERVER}
      if (V_PGI.SAV) and (not V_PGI.ParamSocLast) then
        HShowMessage('0;ERREUR;Paramètre inexistant : '+Nom+';W;O;O;O','','');
{$ENDIF}
    END ;
    Ferme(Q) ;
  end else Result := GetParamSoc (Nom);
end;

function ChercheLibelleGCZoneLibre ( Code : string ) : string;
var Q : TQuery;
  TableCC : string;
begin
  Result := '';
  if GetParamSoc('SO_CPSAISIEAFFAIRE')='COM' then
  begin
    Q := OpenSQL ('SELECT CC_LIBELLE FROM ##DP##.CHOIXCOD WHERE CC_TYPE="ZLI" AND CC_CODE="'+Code+'"',True);
    if not Q.Eof then Result := Q.FindField('CC_LIBELLE').AsString
    else Result := '';
    Ferme (Q);
//  end else Result := RechDom ('GCZONELIBRE',Code,False);
   end else Result := RechDomZoneLibre (Code, False);
end;

procedure TOF_Affaire_Mul_Comp.OnElipsisClickAffairePart1(Sender: TObject);
begin
  if GetParamSoc('SO_CPSAISIEAFFAIRE')='COM' then
    LookupList(TControl(Sender),'','','CC_CODE','CC_LIBELLE','','',True,-1,'SELECT CC_CODE,CC_LIBELLE FROM ##DP##.CHOIXCOD WHERE CC_TYPE="AP1"')
  else
  LookupList(TControl(Sender),'','CHOIXCOD','CC_CODE','CC_LIBELLE','CC_TYPE="AP1"','CC_LIBELLE',True,-1);
end;

procedure TOF_Affaire_Mul_Comp.OnElipsisClickAffairePart2(Sender: TObject);
begin
  if GetParamSoc('SO_CPSAISIEAFFAIRE')='COM' then
    LookupList(TControl(Sender),'','','CC_CODE','CC_LIBELLE','','',True,-1,'SELECT CC_CODE,CC_LIBELLE FROM ##DP##.CHOIXCOD WHERE CC_TYPE="AP2"')
  else
  LookupList(TControl(Sender),'','CHOIXCOD','CC_CODE','CC_LIBELLE','CC_TYPE="AP2"','CC_LIBELLE',True,-1);
end;

procedure TOF_Affaire_Mul_Comp.OnElipsisClickAffairePart3(Sender: TObject);
begin
  if GetParamSoc('SO_CPSAISIEAFFAIRE')='COM' then
    LookupList(TControl(Sender),'','','CC_CODE','CC_LIBELLE','','',True,-1,'SELECT CC_CODE,CC_LIBELLE FROM ##DP##.CHOIXCOD WHERE CC_TYPE="AP3"')
  else
  LookupList(TControl(Sender),'','CHOIXCOD','CC_CODE','CC_LIBELLE','CC_TYPE="AP1"','CC_LIBELLE',True,-1);
end;

Initialization
registerclasses([TOF_Affaire_Mul_Comp]);
end.

