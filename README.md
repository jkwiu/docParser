# docParser
내가 편하게 일하려고 만든 프로그램

1. 사용법
   1. ``docParser.py``가 있는 디렉토리에 파싱할 코드를 위치시킨다.
   2. ``docParser.py``의 ``fileName``의 변수명을 파싱할 코드파일로 변경(``.cs``는 빼도 됨)
   3. ``docParser.py``에서 ``initialStartLine``으로 변환을 시작할 라인의 시작점을 수동으로 정해줘야 한다. 대충 class 밑의 ``주석이 시작되는 부분``부터 하면 된다.