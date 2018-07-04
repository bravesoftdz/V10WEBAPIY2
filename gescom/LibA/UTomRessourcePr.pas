unit UTomRessourcePr;

interface
uses {$IFDEF VER150} variants,{$ENDIF}  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}
     Maineagl,eFiche, eFichList,
{$ELSE}
     FE_Main,Fiche, FichList,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}db,HDB,
{$ENDIF}
     HCtrls,HEnt1,HMsgBox,UTOM,Dialogs,TiersUtil,EntGC,LookUp,
     Dicobtp,HTB97, confidentaffaire, menus,
     utilarticle,UtilPGI,UtilGc, UTOB, UTOF, AGLInit,UtilRessource;


Type
     TOM_RessourcePR = Class (TOM)
       procedure OnNewRecord  ; override ;
       procedure OnArgument (stArgument : String ) ; override ;
       procedure OnChangeField (F : TField)  ; override ;
       procedure OnUpdateRecord  ; override ;
       procedure OnLoadRecord  ; override ;
       procedure OnClose  ; override ;
       procedure CalculFormule;
       procedure FormPRRessource (valeur : string);
       function  GetData(s:Hstring):variant;
       Procedure GereGrisageFamilles;
       Procedure GereGrisageArticle;
     Public
       tsArgsReception : TStringList;
     Private
     Typevalo : string;
     TAUXUNIT : Double;
     TAUXREVIENT : Double;
     TAUXVENTE : Double;
     TAUXFRAISGEN1 : Double;
     TAUXFRAISGEN2: Double;
     COEFMETIER :Double;
     TAUXCHARGEPAT : Double;
     UNITETEMPS : string;
     End;
Procedure AFLanceFiche_RessourcePr(Range,argument:string);

implementation
Uses ParamSoc,Formule,windows, uafo_ressource,  UTOMRessource ;

//var
//   TauxPR :Double;

 (*----------------------------------------------------------------------------*)
//
//
// TOM_RESSOURCEPR
//
//
{Tom Ressource PR}
Procedure TOM_RessourcePR.OnNewRecord;
Var QQ : TQuery ;
    IMax :integer ;
Begin
QQ := nil;
try
If (GetControlText('ARP_RESSOURCE') <> '') then
   QQ:=OpenSQL('SELECT MAX(ARP_RANG) FROM RESSOURCEPR WHERE ARP_RESSOURCE="'+GetField('ARP_RESSOURCE')+'"',TRUE)
Else  begin
   QQ:=OpenSQL('SELECT MAX(ARP_RANG) FROM RESSOURCEPR WHERE ARP_FONCTION="'+GetField('ARP_FONCTION')+'"',TRUE);
   SetField('ARP_RESSOURCE','***'); // mise * dans ressource pour OK appel depuis fiche
   end;

if Not QQ.EOF then Imax:=QQ.Fields[0].AsInteger+1 else iMax:=1;
finally
Ferme(QQ);
If GetControlText('ARP_TYPEVALO') = '' then SetField('ARP_TYPEVALO',Typevalo);
end;

SetField('ARP_RANG',IMax);
End;

procedure TOM_RessourcePR.OnArgument(stArgument : String );
Var
sTmp :string;
MenuTxRevient : TMenuItem;

begin
inherited;  // MCD   pour prise en compte type action
SetControltext('TARP_RESSOURCE', TraduitGA('Ressource'));

if (stArgument = '') then exit;
sTmp := StringReplace(stArgument, ';', chr(VK_RETURN), [rfReplaceAll]);

// On réceptionne les paramètres passés par l'écran appelant, s'il y en a
tsArgsReception := TStringList.Create;

tsArgsReception.Text := sTmp;
if (tsArgsReception.Count = 0) then
   begin
   tsArgsReception.Free;
   tsArgsReception := nil;
   exit;
   end;



// On adapte la configuration de l'écran suivant la provenance de l'appel
if (tsArgsReception.Values['ORIGINE'] = 'ARS') then
   begin
   // Si on vient de l'écran de saisie des ressources
   with tsArgsReception do
   begin
   //If Values['ACTION']= 'CONSULTATION' then TFFiche(ecran)    .TypeAction=taConsult
   //else If Values['ACTION']= 'MODIFICATION' then TFFiche(ecran).TypeAction=taConsult
   TypeValo := Values['TYPEVALO'];
   TAUXUNIT := StrtoFloat(Values['TAUXUNIT']);
   TAUXFRAISGEN1 := StrtoFloat(Values['TAUXFRAISGEN1']);
   TAUXFRAISGEN2 := StrtoFloat(Values['TAUXFRAISGEN2']);
   COEFMETIER := StrtoFloat(Values['COEFMETIER']);
   TAUXCHARGEPAT := StrtoFloat(Values['TAUXCHARGEPAT']);
   if (Values['TAUXREVIENT']<>'') then
        TAUXREVIENT := StrtoFloat(Values['TAUXREVIENT']);
   if (Values['TAUXVENTE']<>'') then
        TAUXVENTE := StrtoFloat(Values['TAUXVENTE']);
   UNITETEMPS := Values['UNITETEMPS'];
   TCheckBox(GetControl('CHBTEMP')).Checked := False;
   end;
   end else begin
       // cas ou on vient depuis la fct. Voir les champs qui seront passés
       // il faudra aussi revoir le menu de sélection des champs possibles pour valorisation
   with tsArgsReception do
      begin
      TypeValo := Values['TYPEVALO'];
      end;
   end;

MenuTxRevient := TMenuItem(GetControl('TAUXREVIENT'));
If MenuTxRevient <> Nil then
    if  (TypeValo='R') then
        MenuTxRevient.Caption := TraduitGA('Prix de revient unitaire')
    else
        MenuTxRevient.Caption := TraduitGA('Prix de vente unitaire');
end;

Procedure TOM_RessourcePR.OnLoadRecord;
Var
BoxArt : TRadioButton;
BoxFam : TRadioButton;
BoxDepAff : TRadioButton;
EvtOnChange:TNotifyEvent;
Str : String;

Begin
TCheckBox(GetControl('CHBTEMP')).Checked := False;

If (GetControlText ('ARP_RESSOURCE') <> '***') then
   Begin
   SetControlVisible ('TARP_FONCTION',False);
   SetControlVisible ('ARP_FONCTION',False);
   End
Else
   Begin
   SetControlVisible ('TARP_RESSOURCE',False);
   SetControlVisible ('ARP_RESSOURCE',False);
   End;

BoxArt := TRadioButton(GetControl('CheckArticle'));
BoxFam := TRadioButton(GetControl('CheckFamille'));
BoxDepAff := TRadioButton(GetControl('CheckDepartement'));
if (getcontroltext('ARP_ARTICLE') <> '') then
   Begin
   If Not (BoxArt.Checked) then
      SetControlProperty ('CheckArticle','Checked',True);
   If (BoxFam.Checked) then
      SetControlProperty ('CheckFamille','Checked',False);
   If (BoxDepAff.Checked) then
      SetControlProperty ('CheckDepartement','Checked',False);
   End Else
if (getcontroltext('ARP_DEPARTEMENT') <> '') then
   Begin
   If Not (BoxDepAff.Checked) then
      SetControlProperty ('CheckDepartement','Checked',True);
   If (BoxArt.Checked) then
      SetControlProperty ('CheckArticle','Checked',False);
   If (BoxFam.Checked) then
      SetControlProperty ('CheckFamille','Checked',False);
   End Else
        Begin
        If Not (BoxFam.Checked) then
            SetControlProperty ('CheckFamille','Checked',True);
        If (BoxArt.Checked) then
            SetControlProperty ('CheckArticle','Checked',False);
        If (BoxDepAff.Checked) then
            SetControlProperty ('CheckDepartement','Checked',False);
        End;

EvtOnChange := THEdit(GetControl('ARP_CODEARTICLE')).OnChange;
THEdit(GetControl('ARP_CODEARTICLE')).OnChange:=nil;
THEdit(GetControl('ARP_CODEARTICLE')).Text := trim(copy(GetField('ARP_ARTICLE'), 1, 18));
THEdit(GetControl('ARP_CODEARTICLE')).OnChange:=EvtOnChange;
if GetControlText('ARP_TYPEVALO') = 'V' then Str := 'Vente' else Str:='Revient' ;
Ecran.Caption := TraduitGA('Personnalisation du Prix de '+Str+' :');
TCheckBox(GetControl('CHBTEMP')).Checked := True;

GereGrisageArticle;
GereGrisageFamilles;
END;



Procedure TOM_RessourcePR.OnUpdateRecord;
Var
BoxArt, BoxDep : TRadioButton;

Begin
// Gestion des erreurs sur le code article
if THEdit(GetControl('ARP_CODEARTICLE')).Text<>'' then
if Not LookupValueExist(GetControl('ARP_CODEARTICLE')) then
Begin LastError:=16; LastErrorMsg:=TraduitGa(TexteMsgRessource[LastError]); SetFocusControl('ARP_CODEARTICLE');Exit; End;

// Gestion du passage du CODEARTICLE en ARTICLE pour stocker
if (THEdit(GetControl('ARP_CODEARTICLE')).Text='') then SetField('ARP_ARTICLE', '')
else SetField('ARP_ARTICLE', CodeArticleUnique(THEdit(GetControl('ARP_CODEARTICLE')).Text,'','','','','')) ;


BoxArt := TRadioButton(GetControl('CheckArticle'));
BoxDep := TRadioButton(GetControl('CheckDepartement'));
If (Not BoxArt.Checked) and (Not BoxDep.Checked) then
   Begin
   setField('ARP_LIBELLE', 'Fam :'+GetField('ARP_FAMILLENIV1')+'/'+GetField('ARP_FAMILLENIV2')+'/'+GetField('ARP_FAMILLENIV3'));
   If ((GetControlText('ARP_FAMILLENIV1') = '') And
       (GetControlText('ARP_FAMILLENIV2') = '') And
       (GetControlText('ARP_FAMILLENIV3') = '')) then Begin LastError:=2 ; LastErrorMsg:=TexteMsgRessource[LastError]; End;
   End
else
   Begin
   If (BoxDep.Checked) then
        Begin
        setField('ARP_LIBELLE', GetField('ARP_DEPARTEMENT'));
        If (GetControlText('ARP_DEPARTEMENT') = '') then Begin LastError:=17 ; LastErrorMsg:=TexteMsgRessource[LastError] ; End;
        end;

   If (BoxArt.Checked) then
        begin
        setField('ARP_LIBELLE', GetField('ARP_ARTICLE'));
        If (GetControlText('ARP_ARTICLE') = '') then Begin LastError:=1 ; LastErrorMsg:=TexteMsgRessource[LastError] ; End;
        end;
   End;


if (tsArgsReception <> nil) then
   begin
   TFFiche(ecran).Retour := '';
   end;
End;

Procedure TOM_RessourcePR.GereGrisageFamilles;
Var BoxArt, BoxDep : TRadioButton;
Begin

   BoxArt := TRadioButton(GetControl('CheckArticle'));
   BoxDep := TRadioButton(GetControl('CheckDepartement'));
   If (BoxArt.Checked) then
      Begin
      SetControlEnabled('ARP_DEPARTEMENT',False);
      SetControlEnabled('ARP_FAMILLENIV1',False);
      SetControlEnabled('ARP_FAMILLENIV2',False);
      SetControlEnabled('ARP_FAMILLENIV3',False);
      SetControlEnabled('ARP_CODEARTICLE',True);
      End
    Else
   If (BoxDep.Checked) then
      Begin
      SetControlEnabled('ARP_DEPARTEMENT',true);
      SetControlEnabled('ARP_FAMILLENIV1',False);
      SetControlEnabled('ARP_FAMILLENIV2',False);
      SetControlEnabled('ARP_FAMILLENIV3',False);
      SetControlEnabled('ARP_CODEARTICLE',False);
      End;
End;

Procedure TOM_RessourcePR.GereGrisageArticle;
Var Box : TRadioButton;
Begin

   Box := TRadioButton(GetControl('CheckFamille'));
   If (Box.Checked) then
      Begin
      SetControlEnabled('ARP_FAMILLENIV1',True);
      SetControlEnabled('ARP_FAMILLENIV2',True);
      SetControlEnabled('ARP_FAMILLENIV3',True);
      SetControlEnabled('ARP_CODEARTICLE',False);
      SetControlEnabled('ARP_DEPARTEMENT',False);
      End;
End;

Procedure TOM_RessourcePR.OnChangeField(F:TField);
//Var Box : TRadioButton;
Begin

(*If (F.FieldName='ARP_CODEARTICLE') then
   Begin
   Box := TRadioButton(GetControl('CheckArticle'));
   If Box.Checked  then
      Begin
      SetControlEnabled('ARP_FAMILLENIV1',False);
      SetControlEnabled('ARP_FAMILLENIV2',False);
      SetControlEnabled('ARP_FAMILLENIV3',False);
      SetControlEnabled('ARP_CODEARTICLE',True);
      SetControlEnabled('ARP_DEPARTEMENT',False);
      End;
   End;
If (F.FieldName='ARP_ARTICLE') then
   Begin
   Box := TRadioButton(GetControl('CheckArticle'));
   If Box.Checked  then
      Begin
      SetControlEnabled('ARP_FAMILLENIV1',False);
      SetControlEnabled('ARP_FAMILLENIV2',False);
      SetControlEnabled('ARP_FAMILLENIV3',False);
      SetControlEnabled('ARP_CODEARTICLE',True);
      SetControlEnabled('ARP_DEPARTEMENT',False);
      End;
   End;
If (F.FieldName='ARP_ARTICLE') then
   Begin
   Box := TRadioButton(GetControl('CheckArticle'));
   If Box.Checked  then
      Begin
      SetControlEnabled('ARP_FAMILLENIV1',False);
      SetControlEnabled('ARP_FAMILLENIV2',False);
      SetControlEnabled('ARP_FAMILLENIV3',False);
      SetControlEnabled('ARP_CODEARTICLE',True);
      SetControlEnabled('ARP_DEPARTEMENT',False);
      End;
   End;
   *)


If (F.FieldName='ARP_FORMULE') then CalculFormule;
End;

procedure TOM_RessourcePR.OnClose  ;
begin
if (tsArgsReception <> nil) then
   begin
   tsArgsReception.Free;
   tsArgsReception := nil;
   end;
end;

procedure TOM_RessourcePR.FormPRRessource (valeur : string);
Var Formule : TEdit;
Begin
Formule := TEdit(GetControl ('ARP_FORMULE'));
Formule.SelText:='['+valeur+']';
SetField('ARP_FORMULE',Formule.Text);
End;

Procedure TOM_RessourcePR.CalculFormule();
Var r, r2 : variant;
    s : Hstring;
Begin
s := '{' + Getfield ('ARP_FORMULE') + '}';
if not TestFormule (s) then
   begin
   ShowMessage ('La formule est incorrecte.');
   SetFocusControl ('ARP_FORMULE');
   exit;
   end;

if ((GetControlText ('ARP_ARTICLE') = '') And
    ((Pos ('TAUXUNITPRES', Getfield ('ARP_FORMULE')) > 0) Or (Pos ('FRAISGENPREST', Getfield ('ARP_FORMULE')) > 0)))
  Then
  SetControlText ('RESULTAT', 'Non Calculable')
Else
    Begin
// Pl le 14/01/03 : format numérique + traitement de l'erreur
    s := '{"#.###,0"' + s + '}';
    r := GFormule (s, GetData, nil, 1);
    try
      if (r <> '') then
        r2 := VarAsType (r, varDouble)
      else
        r := 0;
      SetControlText ('RESULTAT', r);
    except
      SetControlText ('RESULTAT', 'Erreur');
    end;
//    if (VarType(r) = varString) then
//    else
//      SetControlText ('RESULTAT', r);
/////////////////////////////////////////////////////////
    End;

End;

function TOM_RessourcePR.GetData(s:Hstring):variant;
Var Q: TQuery;
dCoeffConvert:double;
Begin
Result :=1;
    // ATTENTION, si modif de cette fct, faire idem dans TAFO_Ressource.DonneesFormulePR (UAFO_RESSOURCE)
s:=AnsiUppercase(s);
if (s= 'TAUXUNITRES') then result := TAUXUNIT else
if (s= 'FRAISGEN1') then result := TAUXFRAISGEN1 else
if (s='FRAISGEN2') then result := TAUXFRAISGEN2 else
if (s='COEFMETIER') then result := COEFMETIER else
if (s='TAUXREVIENT') then result := TAUXREVIENT else
if (s='TAUXVENTE') then result := TAUXVENTE else
if (s='TAUXPATRON') then result := TAUXCHARGEPAT else
   Begin
   // traitement des zones articles
   if (s='TAUXUNITPRES') Or (s='FRAISGENPREST') then
      Begin
       if (GetControlText('ARP_ARTICLE') <> '') then
          Begin
          Q:=OpenSQL ('SELECT GA_PMRP,GA_COEFFG,GA_QUALIFUNITEVTE  FROM ARTICLE WHERE GA_ARTICLE ="'+ GetField ('ARP_ARTICLE')+'"',True);
          if ((Not Q.EOF) And (s='TAUXUNITPRES')) then
            begin
            if (Q.Fields[2].AsString = UNITETEMPS) then
               // si l'article est dans la meme unite que la ressource
               result :=Q.Fields[0].AsFloat
            else
               // Sinon, on doit convertir dans l'unite de la ressource
               begin
               dCoeffConvert := ConversionUnite(Q.Fields[2].AsString, UNITETEMPS, 1);
               if dCoeffConvert=0 then dCoeffConvert:=1;
               result := Q.Fields[0].AsFloat / dCoeffConvert;
               end;
            end
          else
          if ((Not Q.EOF) And (s='FRAISGENPREST')) then result :=Q.Fields[1].AsFloat;
          Ferme(Q);
          End;
      End;
   End;
end;

 Procedure AFLanceFiche_RessourcePr(Range,argument:string);
begin
AGLLanceFiche('AFF','RESSOURCEPR',Range,'',Argument);
end;

Initialization
registerclasses([TOM_RessourcePR]);
end.


