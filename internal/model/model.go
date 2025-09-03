package model

import "time"

type Record struct {
	Noderef        string
	LastName       string
	FirstName      string
	MiddleName     string
	Department     string
	Position       string
	CounterpartyID string
	ModifiedSource time.Time
	Status         string
	ArchiveNodeRef string
}
