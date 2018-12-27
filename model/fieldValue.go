package model

type FieldValue struct {
	Value        string
	Type         string
	ChildrenList []*FieldValue
}
