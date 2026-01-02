package errwrap

import "fmt"

// WrapOp adds a stable operation prefix while preserving the original error.
func WrapOp(op string, err error) error {
	if err == nil {
		return nil
	}
	return fmt.Errorf("%s: %w", op, err)
}
