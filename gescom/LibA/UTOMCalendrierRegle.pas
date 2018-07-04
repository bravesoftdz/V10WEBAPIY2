{***********UNITE*************************************************
Auteur  ...... : CHARRAIX
Créé le ...... : 13/06/2001
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CALENDRIERREGLE (CALENDRIERREGLE)
Mots clefs ... : TOM;CALENDRIERREGLE
*****************************************************************}
Unit UTOMCalendrierRegle ;

Interface

Uses StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}
    Maineagl,
{$ELSE}
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db, HDB, FE_Main,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOM,Dialogs,TiersUtil,EntGC,
      Dicobtp,HTB97,Vierge, UTOB,AGLInit, LicUtil;

Type
  TOM_CALENDRIERREGLE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  private
    TypeHoraire, CodeRes, CodeSalarie, LibNom, Standard, {TypAction , }listeChampGene : String;
    Creat: boolean; //mcd 04/10/02
    procedure ControlDonnee(element : string) ;
    Function VerifChampControlDonnee (champ, lib : string; valeur : double )  : integer;
    //procedure VerifExit (Sender: TObject);
  end ;
Procedure AFLanceFiche_RegleCalendrier(Range,Argument:string);

Implementation

procedure TOM_CALENDRIERREGLE.OnNewRecord ;
var
Q :Tquery;
rang : integer;
rq, param : string;
begin
  Inherited ;
  if (TypeHoraire <> 'STD') then
    begin
     SetField('ACG_STANDCALEN',standard);
     if (TypeHoraire ='RES')  then
        BEGIN
        // SetControlText ('LIBELLE','Calendrier de ' + TraduitGA('la ressource'));
         SetField ('ACG_RESSOURCE', coderes);
        // SetControlText('LIBNOM', LibNom);
         param := 'ACG_RESSOURCE';
        END
     else
        BEGIN
        // SetControlText ('LIBELLE','Calendrier du salarié ') ;
         SetField ('ACG_SALARIE', codesalarie);
        // SetControlText('LIBNOM', LibNom);
         param := 'ACG_SALARIE';
        END;
    rq := 'SELECT MAX(ACG_RANG) FROM CALENDRIERREGLE WHERE ACG_STANDCALEN="'+standard+
        '" AND '+param+'="'+GetField(param)+'"';
    end
  else
    begin
    // SetControlVisible('LIBELLE',False);
    // SetControlVisible('LIBNOM',False);
     rq := 'SELECT MAX(ACG_RANG) FROM CALENDRIERREGLE WHERE ACG_STANDCALEN="'+standard+
           '" AND ACG_RESSOURCE="***" AND ACG_SALARIE="***"';
    end;
  //
  Q := OpenSQL(rq,True);
  If not Q.eof then rang:= Q.Fields[0].asInteger + 1 else rang :=1;
  SetField('ACG_RANG',rang);
  SetField('ACG_LIBELLE','Regle n° '+IntToStr(rang));
  Ferme(Q);
  //
  SetField('ACG_CTRLJOUR','-');
  SetField('ACG_DIMPERMIS','-');
  SetField('ACG_FERIEPERMIS','-');
  SetField('ACG_BLOCDEPHRJ','-');
  SetField('ACG_BLOCDEPHSMN','-');
  SetField('ACG_SEMCONSECLIM1','-');
  SetField('ACG_SEMCONSECLIM2','-');
  SetField('ACG_NBHMAXJL','24');
  SetField('ACG_NBHMAXSMNL','168');
  SetField('ACG_NBMAXHANALERTL','1600');
  SetField('ACG_NBMAXHANBLOCL','2000');
  SetField('ACG_NBDJMAXANL','434');
  //
  SetControlEnabled('BIMPRIMER',False);
  Creat:=true;
end ;

procedure TOM_CALENDRIERREGLE.OnDeleteRecord ;
begin
  Inherited ;
Creat:=False;
end ;

procedure TOM_CALENDRIERREGLE.OnUpdateRecord ;
begin
  ControlDonnee('');
  If GetField('ACG_RESSOURCE') = '' then SetField('ACG_RESSOURCE','***');
  If GetField('ACG_SALARIE') = '' then SetField('ACG_SALARIE','***');
     // mcd 03/04/2002 on interdit saisie d'une date pour même std/calendrier qui existe
     // sur index 2
  If Creat And ExisteSql('SELECT ACG_STANDCALEN from CALENDRIERREGLE WHERE ACG_STANDCALEN="'+
     GetField('ACG_STANDCALEN')+'" AND ACG_RESSOURCE="'+
     GetField('ACG_RESSOURCE')+'" AND ACG_DATEEFFET="'+ DateToSTr(GetField('ACG_DATEEFFET'))+'"')
      then begin
      SetFocusControl('ACG_DATEEFFET');
      LastError:=1;
      lastErrorMsg:='La Date d''effet existe déjà pour ce calendrier';
      exit;
      end;
  SetControlEnabled('BIMPRIMER',True);
  Inherited ;
  Creat:=False;
end ;

procedure TOM_CALENDRIERREGLE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_CALENDRIERREGLE.OnLoadRecord ;
begin
  Inherited ;
  SetControlEnabled('ACG_NBDJMAXAN',GetField('ACG_CTRLJOUR')='X')
end ;

procedure TOM_CALENDRIERREGLE.OnChangeField ( F: TField ) ;
//var txt : string;
begin
  Inherited ;
  //intégré dans le script
  {If (F.FieldName='ACG_CTRLJOUR') then
    begin
      SetControlEnabled('ACG_NBDJMAXAN',GetField('ACG_CTRLJOUR')='X');
    end
  else }

  // vérification des maximas....
  //If (F.FieldName = 'TEMP') and (GetField(F.fieldName) <> '1') then ControlDonnee(GetField(F.fieldName));
end ;

procedure TOM_CALENDRIERREGLE.OnArgument ( S: String ) ;
Var //CC      : THEDIT;
    //Combo   : THValComboBox;
    //CTotal  : THLabel;
    //BtDet,  BtDuplic, BtErase, BtCalreg : TToolBarButton97;
    //BValider : Ttoolbarbutton97;
    Critere,{tempChamp,}Champ, valeur,listechamp  : String;
    X             : integer;
begin
Inherited;
 ListeChampGene := 'ACG_NBHMAXSMN;ACG_NBHMAXJ;ACG_NBMAXHANBLOC;ACG_NBMAXHANALERT;ACG_NBDJMAXAN';
// Récupération des paramètres
// 3 paramétres sont interprétés TYPE: 'STD' 'RES' ou 'SAL'
// CODE: code de la ressource ou salarié LIBELLE : libellé associé
// STANDARD: Code standard calendrier

TypeHoraire := 'STD'; CodeRes := ''; CodeSalarie := ''; LibNom := ''; Standard := '';
Critere:=(Trim(ReadTokenSt(S)));
While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        X:=pos(':',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;
        if (Champ = 'TYPE')   then TypeHoraire := Valeur;
        if (Champ = 'CODESALARIE') then CodeSalarie := Valeur;
        if (Champ = 'CODERES') then CodeRes := Valeur;
        if (Champ = 'LIBELLE') then LibNom := Valeur;
        if (Champ = 'STANDARD')then Standard := Valeur;
        END;
    Critere:=(Trim(ReadTokenSt(S)));
    END;

// Affichage spécifique si calendrier ressource ou salarié
if (TypeHoraire <> 'STD') then
    BEGIN
   if (TypeHoraire ='RES')  then
        BEGIN
         SetControlText ('LIBELLE','Calendrier de ' + TraduitGA('la ressource')+' : '+LibNom);
         SetControlVisible('LIBNOM', False);
         Ecran.Caption := 'Règles du calendrier de '+ TraduitGA('la ressource')+' '+LibNom;
        END
    else
        BEGIN
          SetControlText ('LIBELLE','Calendrier du salarié : '+LibNom) ;
          SetControlVisible('LIBNOM', False);
          Ecran.Caption := 'Règles du calendrier du salarié '+LibNom;
        END;
    END
else
   begin
     SetControlText ('LIBELLE','Calendrier Standard : '+LibNom) ;
     SetControlVisible('LIBNOM', False);
     Ecran.Caption := 'Règles du calendrier standard : '+LibNom;
   end;

//affichage dez zones de paramétrages limites avec mot de passe du jour
if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
 begin
  ListeChamp := ListeChampGene;
   While ListeChamp <>'' do
    begin
     Champ := ReadTokenSt(Listechamp);
     Champ := Champ+'L';
     SetControlVisible(Champ,True);
    end;
   For X:= 1 to 4 do
    begin
     SetControlVisible('Limite'+IntToStr(X),True);
    end;
 end;
 SetControlText('TYPECAL',TypeHoraire);
 // zones non gérées pour l'instant
SetControlVisible('ACG_CTRLJOUR',False);
SetControlVisible('ACG_NBDJMAXAN',False);
SetControlVisible('TACG_NBDJMAXAN',False);

End;


procedure TOM_CALENDRIERREGLE.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_CALENDRIERREGLE.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_CALENDRIERREGLE.ControlDonnee(element :string) ;
var
  ListeChamp, champ, tempChamp{,tempVal} : String;
  Val : Double;
  //Num : ThNumEdit;
begin
If element <>'' then ListeChamp := element else ListeChamp := ListeChampGene;
While listechamp <>'' do
  BEGIN
    Champ := ReadTokenSt(ListeChamp);
    tempchamp := champ+'L';
    Val := GetField(TempChamp);
    if (champ='ACG_NBHMAXSMN') or (champ='ACG_NBHMAXJ')then
       begin
         if VerifChampControlDonnee(champ,TgroupBox(GetControl('GB'+champ)).caption,Val) = 1 then
            exit;
       end
    else
       begin
        If VerifChampControlDonnee(champ,GetControlText('T'+champ),Val) = 1 then exit;
       end;
  END;
end;

Function TOM_CALENDRIERREGLE.VerifChampControlDonnee(champ, lib: string;  valeur: double) : integer;
begin
  if getField(champ) > valeur then
    begin
      PGIInfoAF('La valeur saisie pour le '+lib+' n''est pas cohérente, vous êtes limité à '+FloatToStr(valeur),TitreHalley);
      Result := 1;
      LastError := 1;
      SetFocuscontrol(champ);
    end
  else
    begin
     result := 0;
     LastError := 0;
    end;
end;

Procedure AFLanceFiche_RegleCalendrier(Range,Argument:string);
begin
AGLLanceFiche ('AFF','AFCALENDRIERREGLE',Range,'',Argument);
end;

Initialization
  registerclasses ( [ TOM_CALENDRIERREGLE ] ) ; 
end.
