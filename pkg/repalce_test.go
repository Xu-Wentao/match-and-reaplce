package pkg

import (
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
				str:    "1234567890（测试中邮频道）（24）",
				substr: "（测试*频道）",
			},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			reFirstMatch(tt.args.str, tt.args.substr)
		})
	}
}
