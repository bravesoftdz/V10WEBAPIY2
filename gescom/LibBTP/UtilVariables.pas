unit UtilVariables;

interface
uses lookup,Hctrls;

Function GetVariableRecherche (G_CodeVariable: THCritMaskEdit;
                              stWhere, stRange, stMul : string; bCodeArticle : Boolean = False) : string ;

implementation

// permet d'obtenir le module de recherhce des variables
Function GetVariableRecherche (G_CodeVariable: THCritMaskEdit;
                              stWhere, stRange, stMul : string; bCodeArticle : Boolean = False) : string ;
Var G_Variable : THCritMaskEdit;
begin
  G_Variable := G_CodeVariable;
   LookupList (G_Variable, 'Variables Générales', 'BVARIABLES', 'BVA_CODEVARIABLE', 'BVA_LIBELLE', stWhere,'BVA_CODEVARIABLE', True, 0) ;
if (G_Variable.Text <> '') and (G_Variable.Text <> G_CodeVariable.Text) then G_CodeVariable.Text := G_Variable.Text;
end;


end.
