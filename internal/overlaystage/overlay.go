package overlaystage

type Overlay interface {
	Get(path string) ([]byte, error)
	Set(path string, content []byte) error
	Has(path string) (bool, error)
	ListEntries() ([]string, error)
}

type BaselineChecker interface {
	HasBaseline(path string) (bool, error)
}
