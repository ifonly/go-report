package model

type FieldKey struct {
	Id           string
	Name         string
	Type         string
	TypeName     string
	ShowUnit     bool
	IsRange      bool
	IsMultiLevel bool
	ChildrenList []*FieldKey
}

const (
	TYPE_NUMBER                = "DOUBLE"
	TYPE_DATE                  = "DATE"
	TYPE_PERCENTAGE            = "PERCENTAGE"
	TYPE_FORMULA               = "FORMULA"
	TYPE_UNKNOWN_ENUM_VALUE_10 = "UNKNOWN_ENUM_VALUE_FieldValueType_10"
	SUBMITTER                  = "SUBMITTER"
)
