unit UTomDimMasque;

interface
uses {$IFDEF VER150} variants,{$ENDIF}  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOM, HDimension, Dialogs ,
{$IFDEF EAGLCLIENT}
      eFichList,UTob,
{$ELSE}
      HDB,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}db,FichList,dbCtrls,
{$ENDIF}
      M3FP,EntGC,UtotGC,UtilGc,ParamSoc;

Type
     TOM_Dimmasque = Class (TOM)
       Valeurs : THDimensionItemList ;
       procedure OnChangeField (F : TField)  ; override ;
       procedure OnUpdateRecord ; override ;
       procedure OnAfterUpdateRecord ; override ;
       procedure OnArgument (stArgument : String ) ; override ;
       procedure OnLoadRecord ; override ;
       procedure OnDeleteRecord ; override ;
       procedure OnClose ; override ;
       procedure AfficheDim (stArgument : String; var DimPlus : array of string) ;
       procedure AjusteControle (bAfficheEtab : Boolean ) ;
       procedure SetLastError (Num : integer; ou : string );
     private
       bMultiEtab : boolean ; // Masque multi-établissements
       bFieldTypeModified : Boolean ;  // Un champ GDM_TYPE1..5 a-t-il été modifié ?
       bFieldFermerModified : Boolean; // Le champ GDM_FERMER a été modifié
       DimPlus : array[1..MaxDimension] of string ;
{$IFDEF BTP}
    function CtrlNonVide(Indice, Code: string): boolean;
{$ENDIF}
     END ;

const TexteMessage: array[1..10] of string 	= (
        {1}   'Si une dimension est en taille unique,'+#13+#10+'toutes les dimensions doivent être en taille unique !'
        {2}  ,'Cette position est déjà attribuée à une autre dimension !'
        {3}  ,'Un onglet ne peut pas être utilisé sans ligne ou colonne !'
        {4}  ,'La position de cette dimension doit être renseignée !'
        {5}  ,'Un masque doit comprendre au moins une dimension renseignée !'
        {6}  ,'Des articles utilisent ce masque. Vous ne pouvez pas le supprimer !'
        {7}  ,'D''autres types de ce masque existent et seront également supprimés.'+#13+#10+
              'Confirmez-vous la suppression ?'
        {8}  ,'L''établissement est obligatoire, le type de masque est définit en multi-établissements !'
        {9}  ,'Vous devez d''abord renseigner les valeurs de cette dimension'
        {10} ,'Vous ne pouvez pas supprimer ce masque, il comporte un type de masque.'+#13+#10+
              'Vous pouvez supprimer le masque rattaché au type de masque par défaut'+#13+#10+
              'ou supprimer le type de masque depuis le menu "Types de masques".'
);

implementation

{ TOM_DimMasque  }

procedure TOM_Dimmasque.OnUpdateRecord;
var cpt,cpt2,MaxPos : integer;
    DimVide,DimUnique,DimNonUnique : boolean;
    TypeMasque,Position6,Sql,stTemp : string;
begin
if GetParamSoc('SO_GCARTTYPEMASQUE') then MaxPos:=6 else MaxPos:=5 ;

// Au moins une dimension doit exister
DimVide := True;
for cpt:=1 to 5 do if (GetControlText('GDM_TYPE'+IntToStr(cpt))<>'') then DimVide := False;
if DimVide then BEGIN SetLastError(5,'GDM_TYPE1') ; exit END ;

// Si une dimension est renseignée, alors la position doit être renseignée
for cpt:=1 to 5 do
    if (GetControlText('GDM_TYPE'+IntToStr(cpt))<>'') and
       (GetControlText('GDM_POSITION'+IntToStr(cpt))='')
    then BEGIN SetLastError(4,'GDM_POSITION'+IntToStr(cpt)) ; exit END ;

// Si une position est TAILLE UNIQUE, alors toutes doivent être TAILLE UNIQUE
DimUnique:=False ;
DimNonUnique:=False ;
cpt2:=1 ;
for cpt:=1 to 5 do if (GetControlText('GDM_TYPE'+IntToStr(cpt))<>'') then
    BEGIN
{$IFDEF BTP}
    if not CtrlNonVide (IntTostr(cpt),GetControlText('GDM_TYPE'+IntToStr(cpt))) then
       begin
       SetLastError(9,'GDM_TYPE'+IntToStr(cpt)) ;
       exit
       END;
{$ENDIF}
    If (GetControlText('GDM_POSITION'+IntToStr(cpt))='UNI')
       then BEGIN DimUnique:=True ; cpt2:=cpt END // Une dimension utilisée, en taille unique
       else DimNonUnique:=True ;    // Une dimension trouvée, en taille non unique
    END ;
If (DimUnique) and (DimNonUnique)
    then BEGIN SetLastError(1,'GDM_POSITION'+IntToStr(cpt2)) ; exit END ;

// Ne pas utiliser deux fois la même position
for cpt:=1 to MaxPos do if GetControlText('GDM_POSITION'+IntToStr(cpt))<>'' then
    BEGIN
    for cpt2:=cpt+1 to MaxPos do if GetControlText('GDM_POSITION'+IntToStr(cpt))=GetControlText('GDM_POSITION'+IntToStr(cpt2)) then
        BEGIN SetLastError(2,'GDM_POSITION'+IntToStr(cpt2)) ; exit END ;
    END ;

// position col2 utilisée sans col1
for cpt:=1 to MaxPos do if GetControlText('GDM_POSITION'+IntToStr(cpt))='LI2' then
    BEGIN
    for cpt2:=1 to MaxPos do if GetControlText('GDM_POSITION'+IntToStr(cpt2))='LI1' then break ;
    if cpt2 > MaxPos then Setfield('GDM_POSITION'+IntToStr(cpt),'LI1') ;
    END ;

// position lig2 utilisée sans lig1
for cpt:=1 to MaxPos do if GetControlText('GDM_POSITION'+IntToStr(cpt))='CO2' then
    BEGIN
    for cpt2:=1 to MaxPos do if GetControlText('GDM_POSITION'+IntToStr(cpt2))='CO1' then break ;
    if cpt2 > MaxPos then Setfield('GDM_POSITION'+IntToStr(cpt),'CO1') ;
    END ;

// position onglet utilisée sans col ni lig
for cpt:=1 to MaxPos do if GetControlText('GDM_POSITION'+IntToStr(cpt))='ON1' then
    BEGIN
    for cpt2:=1 to MaxPos do if (GetControlText('GDM_POSITION'+IntToStr(cpt2))='LI1') or (GetControlText('GDM_POSITION'+IntToStr(cpt2))='CO1') then break ;
    if cpt2 > MaxPos then BEGIN SetLastError(3,'GDM_POSITION'+IntToStr(cpt)) ; exit END ;
    END ;

if GetParamSoc('SO_GCARTTYPEMASQUE') then
    BEGIN
    // Contrôle position dépôt
    Position6:=GetControlText('GDM_POSITION6') ;
    TypeMasque:=GetField('GDM_TYPEMASQUE') ;
    if (bMultiEtab) and ((Position6='') or (Position6='AUC'))
       then BEGIN SetLastError(8,'GDM_POSITION6') ; exit END ;

    if ((Position6<>'') and (Position6<>'AUC')) then for cpt:=1 to 5 do
        BEGIN
        // position dépôt sur une position déjà utilisée
        if GetControlText('GDM_POSITION'+IntToStr(cpt))=GetControlText('GDM_POSITION6')
            then BEGIN SetLastError(2,'GDM_POSITION6') ; exit END ;
        END ;

    if bFieldTypeModified then
        BEGIN // Un champ GDM_TYPE1..5 a été modifié -> report pour les autres masques ayant un type différent
        bFieldTypeModified:=False ;
        Sql:='update DIMMASQUE set ' ;
        for cpt:=1 to 5 do
            BEGIN
            stTemp:=IntToStr(cpt) ;
            if cpt>1 then Sql:=Sql+'",' ;
            Sql:=Sql+'GDM_TYPE'+stTemp+'="'+GetField('GDM_TYPE'+stTemp) ;
            END ;
        Sql:=Sql+'" where GDM_MASQUE="'+GetField('GDM_MASQUE')+
                 '" and GDM_TYPEMASQUE<>"'+GetField('GDM_TYPEMASQUE')+'"' ;
        ExecuteSql(Sql) ;
        END ;
    if bFieldFermerModified then
       BEGIN   // le champ GDM_FERMER a été modifié -> report pour les autres masques ayant un type différent
       bFieldFermerModified:=False ;
        Sql:='update DIMMASQUE set GDM_FERMER ="'+GetField('GDM_FERMER') ;
        Sql:=Sql+'" where GDM_MASQUE="'+GetField('GDM_MASQUE')+
                 '" and GDM_TYPEMASQUE<>"'+GetField('GDM_TYPEMASQUE')+'"' ;
        ExecuteSql(Sql) ;
       END;

    END ;

if GetField('GDM_TYPEMASQUE') = '' then SetField('GDM_TYPEMASQUE',VH_GC.BOTypeMasque_Defaut) ;
AvertirTable('GCMASQUEDIM') ;
end ;

procedure TOM_Dimmasque.OnAfterUpdateRecord ;
begin
{$IFDEF EAGLCLIENT}
// AFAIREEAGL
{$ELSE}
SetControlEnabled('GDM_MASQUE',False) ;
SetControlEnabled('GDM_TYPEMASQUE',False) ;
// Création automatique des masques pour chaque type de masque défini
CreerMasqueAuto('',False,GetField('GDM_MASQUE'),DS) ;
DS.Refresh ;
{$ENDIF}
end ;

procedure TOM_Dimmasque.OnLoadRecord ;
var  SQL,typMasq,Position,stTmp : string ;
     cpt,iDim,NbValMax,NbValeurs : integer ;
     Existe : boolean;
     Dim : THDimension;
     QQ : TQuery ;
begin
for cpt:=1 to MaxDimension do DimPlus[cpt]:='' ;
bFieldTypeModified:=False ;  // Un champ GDM_TYPE1..5 a-t-il été modifié ?
bFieldFermerModified:=False ; // Le champ GDM_FERMER a-t-il été modifié ?
// Ajustement taille/visibilité des contrôles de la fiche
typMasq:=GetControlText('GDM_TYPEMASQUE') ;
if DS.State<>DsInsert
    then bMultiEtab:=ExisteSQL('select GMQ_TYPEMASQUE from TYPEMASQUE where GMQ_TYPEMASQUE="'+TypMasq+'" and GMQ_MULTIETAB="X"')
    else bMultiEtab:=True ;
AjusteControle(bMultiEtab) ;
if (typMasq<>'') and (typMasq<>VH_GC.BOTypeMasque_Defaut) then
    BEGIN
{$IFDEF EAGLCLIENT}
    for cpt:=1 to 5 do THValComboBox(GetControl('GDM_POSITION'+IntToStr(cpt))).Plus:='' ;
    THValComboBox(GetControl('GDM_POSITION6')).Plus:='and CO_CODE<>"UNI"' ;
    TCheckBox(GetControl('GDM_FERMER')).Enabled:=False;
{$ELSE}
    for cpt:=1 to 5 do THDBValComboBox(GetControl('GDM_POSITION'+IntToStr(cpt))).Plus:='' ;
    THDBValComboBox(GetControl('GDM_POSITION6')).Plus:='and CO_CODE<>"UNI"' ;
    TDBCheckBox(GetControl('GDM_FERMER')).Enabled:=False;
{$ENDIF}
    END else
    BEGIN
{$IFDEF EAGLCLIENT}
    THValComboBox(GetControl('GDM_TYPEMASQUE')).Plus:='' ;
    for cpt:=1 to 5 do THValComboBox(GetControl('GDM_POSITION'+IntToStr(cpt))).Plus:='and CO_CODE<>"AUC"' ;
    TCheckBox(GetControl('GDM_FERMER')).Enabled:=False;
{$ELSE}
    THDBValComboBox(GetControl('GDM_TYPEMASQUE')).Plus:='' ;
    for cpt:=1 to 5 do THDBValComboBox(GetControl('GDM_POSITION'+IntToStr(cpt))).Plus:='and CO_CODE<>"AUC"' ;
    TDBCheckBox(GetControl('GDM_FERMER')).Enabled:=True;
{$ENDIF}
    END ;

if GetParamSoc('SO_CHARGEDIMDEGRADE') then // ModeDegrade
    BEGIN
    NbValMax:=GetParamSoc('SO_CHARGEDIMMAX') ;
    // Chargement des clauses Plus des combos des grilles de dimensions
    // pour limiter l'utilisation des seuls codes utilisés par l'article.
    for cpt:=1 to MaxDimension do
        BEGIN
        stTmp:=IntToStr(cpt) ; NbValeurs:=0 ;
        Position:=GetControlText('GDM_TYPE'+stTmp) ;
        if Position<>'' then
            BEGIN
            if NbValMax>0 then
                BEGIN
                SQL:='select count(GDI_CODEDIM) as NB from DIMENSION ' +
                     ' where GDI_TYPEDIM="DI'+stTmp +
                     '" and GDI_GRILLEDIM="'+Position+'"' ;
                QQ:=OpenSQL(SQL,True) ;
                if not QQ.EOF then NbValeurs:=QQ.FindField('NB').AsInteger ;
                Ferme(QQ) ;
                END ;

            if NbValeurs>NbValMax then
                BEGIN
                SQL:='select GDI_CODEDIM from DIMENSION ' +
                     'where GDI_TYPEDIM="DI'+stTmp +
                     '" and GDI_GRILLEDIM="'+Position+'"' +
                     ' order by GDI_RANG' ;
                QQ:=OpenSQL(SQL,True) ;
                if not QQ.EOF then
                    BEGIN
                    stTmp:='' ; iDim:=0 ;
                    while (not QQ.EOF) and (iDim<NbValMax) do
                        BEGIN
                        if stTmp<>'' then stTmp:=stTmp+',' ;
                        stTmp:=stTmp+'"'+QQ.FindField('GDI_CODEDIM').AsString+'"' ;
                        inc (iDim) ;
                        QQ.Next ;
                        END ;
                    DimPlus[cpt]:='AND GDI_CODEDIM IN ('+stTmp+') ' ;
                    END ;
                Ferme(QQ) ;
                END ;
            END ;
        END ;
    END ;

AfficheDim (GetControlText('GDM_POSITION1')+';'+GetControlText('GDM_TYPE1')+';'+
            GetControlText('GDM_POSITION2')+';'+GetControlText('GDM_TYPE2')+';'+
            GetControlText('GDM_POSITION3')+';'+GetControlText('GDM_TYPE3')+';'+
            GetControlText('GDM_POSITION4')+';'+GetControlText('GDM_TYPE4')+';'+
            GetControlText('GDM_POSITION5')+';'+GetControlText('GDM_TYPE5')+';'+
            GetControlText('GDM_POSITION6'),DimPlus) ;

// Modification des dimensions impossible si masque utilisé dans article
if (not (DS.State in [dsInsert])) then
    BEGIN
    SQL:='SELECT GA_ARTICLE from ARTICLE where GA_STATUTART="DIM" and GA_DIMMASQUE="'+getField('GDM_MASQUE')+'"' ;
    Existe:=ExisteSQL(SQL) ;
    END
    else existe:=False ;
for cpt:=1 to 5 do SetControlEnabled('GDM_TYPE'+IntToStr(cpt),not Existe) ;

// Déactivation du menu POPUP
Dim:=THDimension(GetControl('VOIRDIM')) ;
if Dim.PopUp<>Nil then BEGIN Dim.PopUp.free ; Dim.PopUp:=Nil END ;
end ;

procedure TOM_Dimmasque.OnChangeField(F: TField);
var cpt : Integer ;
    bMasqueDefaut,bOldMulti : Boolean ;
    Masque,TypeMasque,Position : string ;
    QQ : TQuery ;
begin
bMasqueDefaut:=True ;
if (F.FieldName='GDM_MASQUE') and (GetField('GDM_MASQUE')<>'') and (DS.State in [dsInsert]) then
    BEGIN
    // Si Code + Type défaut existe -> saisie type <> défaut, sinon passage en mode modif.
    QQ:=OpenSQL('select GDM_LIBELLE, GDM_TYPE1,GDM_TYPE2,GDM_TYPE3,GDM_TYPE4,GDM_TYPE5,'+
                'GDM_POSITION1,GDM_POSITION2,GDM_POSITION3,GDM_POSITION4,GDM_POSITION5 '+
                'from DIMMASQUE where GDM_MASQUE="'+GetControlText('GDM_MASQUE')+
                '" and GDM_TYPEMASQUE="'+VH_GC.BOTypeMasque_Defaut+'"',True );
    if not QQ.eof then
        BEGIN
        if GetParamSoc('SO_GCARTTYPEMASQUE') then
            BEGIN
{$IFDEF EAGLCLIENT}
            THValComboBox(GetControl('GDM_TYPEMASQUE')).Plus:='and (GMQ_TYPEMASQUE<>"'+VH_GC.BOTypeMasque_Defaut+'")' ;
{$ELSE}
            THDBValComboBox(GetControl('GDM_TYPEMASQUE')).Plus:='and (GMQ_TYPEMASQUE<>"'+VH_GC.BOTypeMasque_Defaut+'")' ;
{$ENDIF}
            SetControlEnabled('GDM_TYPEMASQUE',True) ;
            SetField('GDM_LIBELLE',QQ.FindField('GDM_LIBELLE').AsString) ;
            for cpt:=1 to 5 do
                BEGIN
                SetField('GDM_TYPE'+IntToStr(cpt),QQ.FindField('GDM_TYPE'+IntToStr(cpt)).AsString) ;
                SetControlEnabled('GDM_TYPE'+IntToStr(cpt),FALSE) ;
                SetField('GDM_POSITION'+IntToStr(cpt),QQ.FindField('GDM_POSITION'+IntToStr(cpt)).AsString) ;
                SetControlEnabled('GDM_POSITION'+IntToStr(cpt),Boolean(GetControlText('GDM_TYPE'+IntToStr(cpt))<>'')) ;
{$IFDEF EAGLCLIENT}
                THValComboBox(GetControl('GDM_POSITION'+IntToStr(cpt))).Plus:='' ;
{$ELSE}
                THDBValComboBox(GetControl('GDM_POSITION'+IntToStr(cpt))).Plus:='' ;
{$ENDIF}
                END ;
            Ferme(QQ) ;
            AfficheDim (GetControlText('GDM_POSITION1')+';'+GetControlText('GDM_TYPE1')+';'+
                        GetControlText('GDM_POSITION2')+';'+GetControlText('GDM_TYPE2')+';'+
                        GetControlText('GDM_POSITION3')+';'+GetControlText('GDM_TYPE3')+';'+
                        GetControlText('GDM_POSITION4')+';'+GetControlText('GDM_TYPE4')+';'+
                        GetControlText('GDM_POSITION5')+';'+GetControlText('GDM_TYPE5')+';'+
                        GetControlText('GDM_POSITION6'),DimPlus);
            // Ajustement taille/visibilité des contrôles de la fiche
            AjusteControle(True) ;
            SetFocusControl('GDM_TYPEMASQUE') ;
            END else
            BEGIN
            Ferme(QQ) ;
            Masque:=GetField('GDM_MASQUE') ;
            TypeMasque:=VH_GC.BOTypeMasque_Defaut ;
{$IFDEF EAGLCLIENT}
// A FAIRE !! ???
{$ELSE}
            DS.Cancel ; // Cancel de l'insert pour se positionner sur l'enreg. existant
{$ENDIF}
            DS.Locate('GDM_MASQUE;GDM_TYPEMASQUE', VarArrayOf([Masque,TypeMasque]), []) ;
            //DS.Edit ;
            END ;
        END else
        BEGIN
        Ferme(QQ) ;
{$IFDEF EAGLCLIENT}
        THValComboBox(GetControl('GDM_TYPEMASQUE')).Plus:='' ;
{$ELSE}
        THDBValComboBox(GetControl('GDM_TYPEMASQUE')).Plus:='' ;
{$ENDIF}
        SetField('GDM_TYPEMASQUE',VH_GC.BOTypeMasque_Defaut) ;
        SetControlEnabled('GDM_TYPEMASQUE',False) ;
        DS.Edit ;
        // Ajustement taille/visibilité des contrôles de la fiche
        AjusteControle(False) ;
        for cpt:=1 to 5 do THDBValComboBox(GetControl('GDM_POSITION'+IntToStr(cpt))).Plus:='and CO_CODE<>"AUC"' ;
        SetFocusControl('GDM_LIBELLE') ;
        END ;
    END ;
if (F.FieldName='GDM_TYPEMASQUE') and (GetField('GDM_TYPEMASQUE')<>'') and (DS.State in [dsInsert]) then
    BEGIN
    // Si existe déjà, passage en mode modif.
    Masque:=GetField('GDM_MASQUE') ;
    TypeMasque:=GetField('GDM_TYPEMASQUE') ;
    if ExisteSQL('select GDM_MASQUE from DIMMASQUE where GDM_MASQUE="'+Masque+'" and GDM_TYPEMASQUE="'+TypeMasque+'"') then
        BEGIN
{$IFDEF EAGLCLIENT}
// A FAIRE !! ???
{$ELSE}
        DS.Cancel ; // Cancel de l'insert pour se positionner sur l'enreg. existant
{$ENDIF}
        DS.Locate('GDM_MASQUE;GDM_TYPEMASQUE', VarArrayOf([Masque,TypeMasque]), []) ;
        END ;
    bMasqueDefaut:=Boolean(TypeMasque=VH_GC.BOTypeMasque_Defaut) ;
    if not bMasqueDefaut then
        BEGIN
        for cpt:=1 to 5 do
            BEGIN
            SetControlEnabled('GDM_TYPE'+IntToStr(cpt),FALSE) ;
            SetControlEnabled('GDM_POSITION'+IntToStr(cpt),Boolean(GetControlText('GDM_TYPE'+IntToStr(cpt))<>'')) ;
{$IFDEF EAGLCLIENT}
            THValComboBox(GetControl('GDM_POSITION'+IntToStr(cpt))).Plus:='' ;
{$ELSE}
            THDBValComboBox(GetControl('GDM_POSITION'+IntToStr(cpt))).Plus:='' ;
{$ENDIF}
            END ;
        END
{$IFDEF EAGLCLIENT}
        else for cpt:=1 to 5 do THValComboBox(GetControl('GDM_POSITION'+IntToStr(cpt))).Plus:='and CO_CODE<>"AUC"' ;
{$ELSE}
        else for cpt:=1 to 5 do THDBValComboBox(GetControl('GDM_POSITION'+IntToStr(cpt))).Plus:='and CO_CODE<>"AUC"' ;
{$ENDIF}
    bOldMulti:=bMultiEtab ;
    bMultiEtab:=ExisteSQL('select GMQ_TYPEMASQUE from TYPEMASQUE where GMQ_TYPEMASQUE="'+TypeMasque+'" and GMQ_MULTIETAB="X"') ;
    if (bOldMulti) and (not bMultiEtab) and (GetField('GDM_POSITION6')<>'') then SetField('GDM_POSITION6','') ;
    AjusteControle(bMultiEtab) ;
    END ;

if bMasqueDefaut then
    BEGIN
    bFieldFermerModified:=(Ds.State=dsEdit) ;
    // Si le type est vide alors pas d'affectation possible de ligne-colonne-onglet
    for cpt:=1 to 5 do if F.FieldName='GDM_TYPE'+IntToStr(cpt) then
        BEGIN
        bFieldTypeModified:=(Ds.State=dsEdit) ;
        Position:='GDM_POSITION'+IntToStr(cpt) ;
        if GetControlText('GDM_TYPE'+IntToStr(cpt))='' then
            BEGIN
            Setfield(Position,'') ;
            SetControlEnabled(Position,False) ;
            END else
            begin
{$IFDEF BTP}
            if not CtrlNonVide (IntTostr(cpt),GetControlText('GDM_TYPE'+IntToStr(cpt))) then
               begin
               SetLastError(9,'GDM_TYPE'+IntToStr(cpt)) ;
               exit
               END;
{$ENDIF}
            SetControlEnabled(Position,True) ;
            end;
        END ;
    END ;
end ;

procedure TOM_Dimmasque.OnArgument(stArgument: String);
Var
    cpt, iCol : integer ;
    bVisible : Boolean ;
    stDim : String ;
begin
Inherited OnArgument (StArgument) ;

if not GetParamSoc('SO_GCARTTYPEMASQUE') then // Seul le masque par défaut existe !
    BEGIN
    SetControlVisible('TGDM_TYPEMASQUE',False) ;
    SetControlVisible('GDM_TYPEMASQUE',False) ;
    SetControlVisible('LDEPOT',False) ;
    SetControlVisible('GDM_POSITION6',False) ;
    SetControlProperty('GB_CHOIXDIM','HEIGHT',165) ;
    SetControlProperty('VOIRDIM','HEIGHT',190) ;
    END ;

if not (ctxMode in V_PGI.PGIContexte) then
{$IFDEF EAGLCLIENT}
    for cpt:=1 to 5 do THValComboBox(GetControl('GDM_POSITION'+IntToStr(cpt))).Plus:='and CO_CODE<>"AUC"' ;
{$ELSE}
    for cpt:=1 to 5 do THDBValComboBox(GetControl('GDM_POSITION'+IntToStr(cpt))).Plus:='and CO_CODE<>"AUC"' ;
{$ENDIF}

// Mise en forme des libellés des dimensions
for iCol:=1 to MaxDimension do
    begin
    stDim:=IntToStr(iCol) ;
    bVisible:=ChangeLibre2('DIMENSION'+stDim,Ecran);
    SetControlProperty('GDM_POSITION'+stDim,'Visible',bVisible) ;
    end ;

end ;

procedure TOM_Dimmasque.OnDeleteRecord  ;
var  SQL : string ;
begin
if GetParamSoc('SO_GCARTTYPEMASQUE') then
    BEGIN
    if GetField('GDM_TYPEMASQUE')=VH_GC.BOTypeMasque_Defaut then
        BEGIN
        // Controle non utilisation dans la définition des articles
        SQL:='select GA_ARTICLE from ARTICLE where GA_STATUTART="GEN" and GA_DIMMASQUE="'+
             GetField('GDM_MASQUE')+'"' ;
        if ExisteSQL (SQL) then BEGIN SetLastError(6,'') ; exit END ;

        // Controle non existence d'autres types de masques : si existe, confirmation suppression de tous
        SQL:='select GDM_MASQUE from DIMMASQUE where GDM_MASQUE="'+GetField('GDM_MASQUE')+
             '" and GDM_TYPEMASQUE<>"'+VH_GC.BOTypeMasque_Defaut+'"' ;
        if ExisteSQL (SQL) then
            BEGIN
            if PGIAsk(TexteMessage[7], Ecran.Caption) = mrYes then
              BEGIN
              ExecuteSQL('delete from DIMMASQUE where GDM_MASQUE="'+GetField('GDM_MASQUE')+
                         '" and GDM_TYPEMASQUE<>"'+VH_GC.BOTypeMasque_Defaut+'"') ;
              DS.Refresh ;
              END
              else BEGIN LastError:=-1 ; exit END ;
            END ;
        END
        else // DCA - FQ MODE 10726 - Pas de suppression d'un masque comportant un type de masque autre que celui par défaut.
        BEGIN SetLastError(10,'') ; exit END ;
    END else
    BEGIN
    SQL:='select GA_ARTICLE from ARTICLE where GA_STATUTART="GEN" and GA_DIMMASQUE="'+
         GetField('GDM_MASQUE')+'"' ;
    if ExisteSQL (SQL) then BEGIN SetLastError(6,'') ; exit END ;
    END ;
end ;

procedure TOM_Dimmasque.OnClose ;
begin
Valeurs.Free ;
end ;

procedure TOM_Dimmasque.SetLastError (Num : integer; ou : string );
begin
LastError:=Num ;
LastErrorMsg:=TexteMessage[LastError] ;
if ou<>'' then SetFocusControl(ou) ;
end ;

procedure TOM_Dimmasque.AfficheDim (stArgument: String; var DimPlus : array of string) ;
var criteres : string;
    Dim : THDimension;
    St,St2,Plus,DataType,Titre : String ;
    i,MaxDim : Integer ;
    TitreOngl,TitreLig1,TitreLig2,TitreCol1,TitreCol2 : string ;
    DiOn,DiL1,DiL2,DiC1,DiC2 : THValComboBox ;
    Par : TForm ;
    bdim : boolean;
begin
bdim:=False ;                   // masque sans aucune dimension
criteres:=stArgument ;
Dim:=THDimension (Getcontrol ('VOIRDIM')) ;
if Dim=Nil then exit ;
DiON:=Nil ; DiL1:=Nil ; DiL2:=Nil ; DiC1:=Nil ; DiC2:=Nil ;
Valeurs.Free ;
Valeurs:=THDimensionItemList.create ;

Dim.FreeCombo:=TRUE ;
Dim.Active:=FALSE ;
Dim.FreeCombo:=FALSE ;
//Dim.TypeDonnee:=otReel ;
Par:=TForm(Dim.OWner) ;
if GetParamSoc('SO_GCARTTYPEMASQUE') then MaxDim:=MaxDimension+1 else MaxDim:=MaxDimension ;
for i:=1 to MaxDim do
    BEGIN
    St:=uppercase(Trim(ReadTokenSt(Criteres))) ;
    St2:=uppercase(Trim(ReadTokenSt(Criteres))) ;
    if (GetParamSoc('SO_GCARTTYPEMASQUE')) and (i=MaxDim) then
        BEGIN
        //Plus:='ET_ETABLISSEMENT=ET_ETABLISSEMENT ORDER BY ET_ETABLISSEMENT' ;
        Plus:='GDE_DEPOT=GDE_DEPOT ORDER BY GDE_DEPOT' ;
        DataType:='GCDEPOT' ; // 'TTETABABR'
        if VH_GC.GCMultiDepots
          then Titre := TraduireMemoire ( 'Dépôt' )
          else Titre := TraduireMemoire ( 'Etablissement' ) ;
        END else
        BEGIN
        // DimPlus[i-1] car indice tableau passé en paramètre commence à 0 !!
        Plus:='GDI_TYPEDIM="DI'+IntToStr(i)+'" AND GDI_GRILLEDIM="'+St2+'" '+DimPlus[i-1]+'ORDER BY GDI_RANG' ;
        //Plus:='GDI_TYPEDIM="DI'+IntToStr(i)+'" AND GDI_GRILLEDIM="'+St2+'" ORDER BY GDI_RANG' ;
        DataType:='GCDIMENSION' ;
        Titre:=RechDom('GCCATEGORIEDIM','DI'+IntToStr(i),FALSE) ;
        END ;
    if St='LI1' then
        BEGIN
        DiL1:=THValComboBox.Create(Par) ; DiL1.Parent:=Par ; DiL1.Visible:=FALSE ;
        DiL1.Plus:=Plus ;
        DiL1.DataType:=DataType ;
        TitreLig1:=Titre ;
        bdim:=True ;
        END
    else if St='LI2' then
        BEGIN
        DiL2:=THValComboBox.Create(Par) ; DiL2.Parent:=Par ; DiL2.Visible:=FALSE ;
        DiL2.Plus:=Plus ;
        DiL2.DataType:=DataType ;
        TitreLig2:=Titre ;
        bdim:=True ;
        END
    else if St='CO1' then
        BEGIN
        DiC1:=THValComboBox.Create(Par) ; DiC1.Parent:=Par ; DiC1.Visible:=FALSE ;
        DiC1.Plus:=Plus ;
        DiC1.DataType:=DataType ;
        TitreCol1:=Titre ;
        bdim:=True ;
        END
    else if St='CO2' then
        BEGIN
        DiC2:=THValComboBox.Create(Par) ; DiC2.Parent:=Par ; DiC2.Visible:=FALSE ;
        DiC2.Plus:=Plus ;
        DiC2.DataType:=DataType ;
        TitreCol2:=Titre ;
        bdim:=True ;
        END
    else if St='ON1' then
        BEGIN
        DiOn:=THValComboBox.Create(Par) ; DiOn.Parent:=Par ; DiOn.Visible:=FALSE ;
        DiOn.Plus:=Plus ;
        DiOn.DataType:=DataType ;
        TitreOngl:=Titre ;
        bdim:=True ;
        END ;
    END ;
if not bdim
   then Dim.active:=False
   else Dim.InitDimension(TitreOngl,TitreLig1,TitreLig2,TitreCol1,TitreCol2,DiOn,DiL1,DiL2,DiC1,DiC2,Valeurs) ;
end ;

procedure TOM_Dimmasque.AjusteControle (bAfficheEtab : Boolean ) ;
begin
if not GetParamSoc('SO_GCARTTYPEMASQUE') then exit ;
SetControlEnabled('LDEPOT',bAfficheEtab) ;
SetControlEnabled('GDM_POSITION6',bAfficheEtab) ;
end ;

procedure AGLAfficheDim( parms: array of variant; nb: integer );
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFicheListe) then OM:=TFFicheListe(F).OM else exit ;
if (OM is TOM_Dimmasque) then TOM_Dimmasque(OM).AfficheDim(Parms[1],TOM_Dimmasque(OM).DimPlus) ;
end ;

{$IFDEF BTP}
function TOM_Dimmasque.CtrlNonVide (Indice,Code : string): boolean ;
begin
result:=ExisteSQL('SELECT GDI_TYPEDIM from DIMENSION where GDI_TYPEDIM="DI'+indice+'" AND GDI_GRILLEDIM="'+Code+'"');
end;
{$ENDIF}

Initialization
registerclasses([TOM_DimMasque]) ;
RegisterAglProc('AfficheDim',TRUE,1,AGLAfficheDim) ;
end.

