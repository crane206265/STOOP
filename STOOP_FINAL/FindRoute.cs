using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.IO.Pipelines;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices.Swift;
using System.Text;
using System.Threading.Tasks;
using RouteFinder;

namespace RouteFinder
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
        }
    }

    public class K_Route
    {
        public int k;
        public K_Route(int k_val)
        {
            k = k_val; // *버려도 됨
        }

        public List<int> FindRoute(List<int> idxs, List<double> RAs, List<double> Decs, List<double> ObsTimes)
        {
            double currenttime = 0;
            int N = idxs.Count;
            List<int> SolIdxs = [];

            int epoch = 0; //epoch for break
            while (true)
            {
                // SubRoute를 탐색할 subgroup 생성
                List<int> subgroup_idxs = [];
                List<double> subgroup_RAs = [];
                List<double> subgroup_Decs = [];
                List<bool> candidates_temp = TemporalCandidates(RAs, Decs, currenttime, SolIdxs);
                for (int i = 0; i < N; i++)
                {
                    if (candidates_temp[i])
                    {
                        subgroup_idxs.Add(idxs[i]);
                        subgroup_RAs.Add(RAs[i]);
                        subgroup_Decs.Add(Decs[i]);
                    }
                }

                // SubRoute 계산 및 시간 업데이트
                (List<int> subroute_temp, double subroutelength) = FindSubRoute(subgroup_idxs, subgroup_RAs, subgroup_Decs, currenttime);
                SolIdxs.AddRange(subroute_temp);
                currenttime += subroutelength; // 움직이는데 걸리는 시간
                for (int i = 0; i < subroute_temp.Count; i++) currenttime += ObsTimes[i]; //총 노출시간

                if (SolIdxs.Count == N) break; // 경로에 모든 천체가 포함되었을 경우 중지
                // ----------------------------------------------------
                // 무한 루프 방지
                // ----------------------------------------------------
                if (epoch == 100) break;

                epoch += 1;
            }

            return SolIdxs;
        }

        private List<bool> TemporalCandidates(List<double> RAs, List<double> Decs, double time, List<int> DoneIdxs)
        {
            // [주어신 시각에 관측 가능한 천체들 후보 리스트(bool) 반환]
            // RAs, Decs : 모든 천체들의 좌표
            // time : 현재 시간
            // DoneIdxs : 이미 관측해서 고려 안해도 될 놈들

            List<bool> Candidates = [];

            for (int i = 0; i < RAs.Count; i++)
            {
                Candidates[i] = true;
                Candidates[i] = Candidates[i] && RiseObservable(RAs[i], Decs[i], time);
                Candidates[i] = Candidates[i] && CloudObservable(RAs[i], Decs[i], time);
                Candidates[i] = Candidates[i] && ObstacleObservable(RAs[i], Decs[i], time);

                bool DoneBool = true;
                for (int j = 0; j < DoneIdxs.Count; j++) if (DoneIdxs[j] == i) DoneBool = false; // 이미 관측한 천체인지 검사
                Candidates[i] = Candidates[i] && DoneBool;
            }
            return Candidates;
        }

        private (List<int>, double) FindSubRoute(List<int> idxs, List<double> RAs, List<double> Decs, double starttime)
        {
            // [주어진 후보군 내 최적의 경로를 Brute-Force(노가다)로 찾기]
            // 시간각이 가장 뒤쪽에 있는 천체는 도착지점으로 고정.
            // idxs : subgroup(subroute를 탐색할 천체들의 인덱스)
            // RAs, Decs : subgroup의 좌표들
            // starttime : subgroup 탐색 시작 시간

            int N = idxs.Count;
            double[] HAs = new double[N];
            double[,] distances = new double[N, N];

            // 시간각 리스트 계산 & 마지막 천체 (시간각이 가장 뒤쪽에 있는 천체) 탐색
            double min_HA = 999999.99;
            int min_HA_idx = 0;
            for (int i = 0; i < idxs.Count; i++)
            {
                HAs[i] = HA_cal(RAs[i], Decs[i], starttime);
                if (HAs[i] < min_HA)
                {
                    min_HA = HAs[i];
                    min_HA_idx = i;
                }
            }

            // 거리-인접 행렬 계산
            for (int i = 0; i < idxs.Count; i++)
            {
                for (int j = 0; j < i; j++)
                {
                    distances[i, j] = MovingTime(RAs[i], Decs[i], RAs[j], Decs[j]);
                    distances[j, i] = MovingTime(RAs[j], Decs[j], RAs[i], Decs[i]);
                }
            }

            List<int> SubRouteObjects = []; // Brute-Force 위한 순열을 할 대상들 (마지막 도착지 제외 전부)
            for (int i = 0; i < idxs.Count; i++)
            {
                if (i == min_HA_idx) continue;
                SubRouteObjects.Add(i);
            }

            // Brute-Force
            double[] pathLengths = new double[Factorial(N - 1)];

            var perms = new List<List<int>>();
            // 순열 ([1, 2, 3], [1, 3, 2], ...) : int의 list
            // perms : 그 순열들을 저장한 리스트
            Generate(SubRouteObjects, 0, perms); // Brute-Force를 위한 순열 리스트 생성

            // 각 순열별(경로별) pathlength 계산 & 최단 경로 탐색
            double min_pathlength = 999999999.9;
            int min_pathlength_idx = 0;
            double pathlength_temp = 0;
            for (int i = 0; i < perms.Count; i++)
            {
                pathlength_temp = 0;
                for (int j = 0; j < perms[i].Count - 1; j++)
                {
                    pathlength_temp += distances[perms[i][j], perms[i][j + 1]]; //생성한 순열을 관측 순서로 생각
                    // perms : 순열들의 리스트 --> perms[i] : 순열
                    // perms[i][j] : i번째 순열의 j번째 요소
                }
                pathlength_temp += distances[perms[i][-1], min_HA_idx];
                pathLengths[i] = pathlength_temp;

                if (pathlength_temp < min_pathlength)
                {
                    min_pathlength = pathlength_temp;
                    min_pathlength_idx = i;
                }
            }

            List<int> sol = []; // subroute의 해답
            sol.AddRange(perms[min_pathlength_idx]);
            sol.Add(min_HA_idx); //마지막 천체를 subroute에 추가해주기
            return (sol, min_pathlength);
        }

        private void Generate<T>(List<T> list, int k, List<List<T>> result)
        {
            if (k == list.Count)
                result.Add(new List<T>(list));
            else
            {
                for (int i = k; i < list.Count; i++)
                {
                    Swap(list, k, i);
                    Generate(list, k + 1, result);
                    Swap(list, k, i); // backtrack
                }
            }
        }

        private void Swap<T>(List<T> list, int i, int j)
        {
            T temp = list[i];
            list[i] = list[j];
            list[j] = temp;
        }

        private int Factorial(int n)
        {
            // 팩토리얼 계산
            if (n < 0) throw new ArgumentException("음수는 계산할 수 없습니다.");
            if (n > 13) throw new ArgumentException("너무 큰 수입니다."); // 수정해도 됨
            int result = 1;
            for (int i = 2; i <= n; i++) result *= n;
            return result;
        }


        private bool RiseObservable(double RA, double Dec, double time)
        {
            // 천체가 떴는가
            return true;
        }

        private bool CloudObservable(double RA, double Dec, double time)
        {
            // 천체가 구름에 가리지 않았는가
            return true;
        }

        private bool ObstacleObservable(double RA, double Dec, double time)
        {
            // 천체가 지형지물 / 장애물에 가리지 않았는가
            return true;
        }

        private double HA_cal(double RA, double Dec, double time)
        {
            return 0.0;
        }

        private double MovingTime(double RA1, double Dec1, double RA2, double Dec2)
        {
            return 0.0;
        }
    }
}
