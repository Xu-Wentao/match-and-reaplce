package pkg

import (
	"fmt"
	"testing"
)

func Test_re(t *testing.T) {
	type args struct {
		str    string
		substr string
	}
	tests := []struct {
		name string
		args args
	}{
		{
			name: "test1",
			args: args{
				str:    "中央1套(综合频道)(24)",
				substr: "(24)",
			},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			fmt.Println(reFirstMatch(tt.args.str, tt.args.substr))
		})
	}
}
