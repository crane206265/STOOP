using ASCOM.DeviceInterface;
using ASCOM.DriverAccess;
//using ASCOM.Interfaces;
using AstroAlgo.Basic;
using AstroAlgo.Models;
using Astronomy;
using OxyPlot;
using OxyPlot.Annotations;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot.WindowsForms;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace STOOP_FINAL
{
    public partial class Form1 : Form
    {
        // 프로그램 시작시 실행
        public Form1()
        {
            InitializeComponent();
            UpdateTelescopeStatus();
            InitializeSearchResults();
            DisplayInitialValues(); // 초기값 표시
            GenerateSkyPlot();
            this.FormClosing += Form1_FormClosing; // FormClosing 이벤트 등록
            finder = new OptimizedRouteFinder(this, telescope);
        }

        private OptimizedRouteFinder finder;
        // 인스턴스 생성


        // 이벤트 핸들러 중복 등록 방지
        private bool isEventHandlerRegistered = false;


        // telescope 선언
        private ASCOM.DriverAccess.Telescope telescope;


        // 높이 (그래프 그릴 때 사용)
        private double hight;


        // 딥스카이 객체 정보를 담을 클래스
        public class DeepSkyObject
        {
            public string Name { get; set; }
            public string Type { get; set; }
            public string RA { get; set; }        // 적경
            public string Dec { get; set; }       // 적위
            public double Magnitude { get; set; }  // 등급 (V-Mag)
            public string Constellation { get; set; } // 별자리
        }


        // 별 객체 정보를 담을 클래스
        public class StarObject
        {
            public string Names { get; set; }     // 여러 이름을 합친 문자열
            public string RA { get; set; }
            public string Dec { get; set; }
            public double Magnitude { get; set; }
            public string Constellation { get; set; } // 추가
        }


        // 관측대상 클래스
        public class ObservationTarget
        {
            public string Name { get; set; }
            public double RA { get; set; }
            public double Dec { get; set; }
            public double ObsTime { get; set; } // 분 단위
            public override string ToString()
            {
                // 관측대상 정보를 보기 좋게 한 줄에 구분
                return $"{Name} | RA: {RA:F3} | Dec: {Dec:F3} | 관측 시간: {ObsTime}분";
            }
        }


        // 딥스카이 객체 리스트
        private List<ObservationTarget> observationTargets = new List<ObservationTarget>();


        // 초기값 위도/경도/고도 상수
        private double DefaultLatitude = 36.0;      // 도 단위
        private double DefaultLongitude = 128.0;    // 도 단위
        private double DefaultAltitude = 0.0;       // 미터 단위


        // 클래스 필드에 타이머 추가
        private Timer GenerateSkyPlotTimer;
        private Timer GenerateHorizonPlotTimer;


        private List<double> obstacleAzList = new List<double>();
        private List<double> obstacleAltList = new List<double>();
        private bool isObstacleEditMode = false;

        //============================================================================//


        // 위도, 경도, 고도 라벨 초기값 표시
        private void DisplayInitialValues()
        {
            textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;
            textBox6.ReadOnly = false;

            button25.Enabled = true;

            label3.Text = "관측지 위도(초기값: 36도)";
            label4.Text = "관측지 경도(초기값: 128도)";
            label12.Text = "관측지 고도(초기값: 0m)";

            textBox2.Text = "연결 필요";
            textBox3.Text = $"{DefaultLatitude}";
            textBox4.Text = $"{DefaultLongitude}";
            textBox6.Text = $"{DefaultAltitude}";
        }


        // 관측지 위도, 경도 가져오기 (가대 연결되어 있으면 가대에서, 아니면 기본값)
        public (double lat, double lon) GetObservingSite()
        {
            double lat = DefaultLatitude;
            double lon = DefaultLongitude;

            try
            {
                // 텍스트박스에 수동 입력된 값이 있으면 우선 사용
                if (double.TryParse(textBox3.Text, out double manualLat))
                    lat = manualLat;
                if (double.TryParse(textBox4.Text, out double manualLon))
                    lon = manualLon;

                // 가대 연결 시 가대의 값을 우선 적용
                if (telescope != null && telescope.Connected)
                {
                    lat = telescope.SiteLatitude;
                    lon = telescope.SiteLongitude;
                }
            }
            catch
            {
                // 예외 발생 시 기본값 유지
            }

            return (lat, lon);
        }


        // 지평좌표계 기준 180도 및 2차원 하늘 뷰어 (방위각 - 고도)
        private void GenerateSkyPlot()
        {
            var site = GetObservingSite();

            // OxyPlot 극좌표 그래프 생성 (방위각-고도)
            var plotModel = new OxyPlot.PlotModel
            {
                PlotAreaBorderColor = OxyPlot.OxyColors.Transparent,
                Background = OxyPlot.OxyColor.FromRgb(45, 55, 65),
                TextColor = OxyPlot.OxyColors.White,
            };

            // -------------------- 자오선(남쪽, 방위각 180도) 굵은 선 추가 --------------------

            double lat = site.lat; // 서울 기본값

            // 자오선: 천정(고도 90, 방위각 180) ~ 수평선(고도 0, 방위각 180) ~ 천구의 북극(고도 = 위도, 방위각 0)
            var meridianLine = new OxyPlot.Series.LineSeries
            {
                Color = OxyPlot.OxyColors.Red,
                StrokeThickness = 2,
                LineStyle = OxyPlot.LineStyle.Solid,
                Title = "자오선"
            };

            // 수평선(가장자리, 남쪽)
            meridianLine.Points.Add(new OxyPlot.DataPoint(0, 180));

            // 천정(중심, 남쪽)
            meridianLine.Points.Add(new OxyPlot.DataPoint(90, 180));

            // 천구의 북극(고도=위도, 방위각=0)
            meridianLine.Points.Add(new OxyPlot.DataPoint(90 - lat, 0));
            plotModel.Series.Add(meridianLine);

            // 천구의 북극을 점으로 표시
            var northPolePoint = new OxyPlot.Annotations.PointAnnotation
            {
                X = 90 - lat, // 고도=위도
                Y = 0,   // 방위각=0 (북쪽)
                Shape = OxyPlot.MarkerType.Circle,
                Size = 8,
                Fill = OxyPlot.OxyColors.DeepSkyBlue,
                Stroke = OxyPlot.OxyColors.White,
                StrokeThickness = 2,
                ToolTip = "천구의 북극"
            };
            plotModel.Annotations.Add(northPolePoint);

            // 방위각 축 (0~360도, 북쪽=0, 시계방향)
            var angleAxis = new OxyPlot.Axes.AngleAxis
            {
                Minimum = 0,
                Maximum = 360,
                MajorStep = 45,
                MinorStep = 15,
                StartAngle = 90,
                EndAngle = 450,
                FormatAsFractions = false,
                MajorGridlineStyle = OxyPlot.LineStyle.Solid,
                MinorGridlineStyle = OxyPlot.LineStyle.Dot,
                LabelFormatter = v =>
                {
                    switch ((int)v)
                    {
                        case 0: return "N";
                        case 90: return "E";
                        case 180: return "S";
                        case 270: return "W";
                        default: return v.ToString("0") + "°";
                    }
                },
                TicklineColor = OxyPlot.OxyColors.White,
                AxislineColor = OxyPlot.OxyColors.White,
                TextColor = OxyPlot.OxyColors.White,
                MajorGridlineColor = OxyPlot.OxyColors.White,
                MinorGridlineColor = OxyPlot.OxyColors.White
            };
            plotModel.Axes.Add(angleAxis);

            // 고도 축 (0~90도, 0=가장자리, 90=중심)
            var radiusAxis = new OxyPlot.Axes.MagnitudeAxis
            {
                Minimum = 0,
                Maximum = 90,
                MajorStep = 30,
                MinorStep = 30,
                StartPosition = 0,
                EndPosition = 1,
                LabelFormatter = v => (90 - v).ToString("0") + "°",
                TicklineColor = OxyPlot.OxyColors.White,
                AxislineColor = OxyPlot.OxyColors.White,
                TextColor = OxyPlot.OxyColors.White,
                MajorGridlineColor = OxyPlot.OxyColors.White,
                MinorGridlineColor = OxyPlot.OxyColors.White
            };
            plotModel.Axes.Add(radiusAxis);

            // 외부(관측 불가 영역) 폴리곤
            // 1. 전체 원(고도=0, 방위각 0~360)을 시계방향으로 따라감
            var outerPoints = new List<OxyPlot.DataPoint>();
            for (int az = 0; az <= 360; az += 5)
                outerPoints.Add(new OxyPlot.DataPoint(90, az));

            var polygonOuter = new OxyPlot.Annotations.PolygonAnnotation
            {
                Fill = OxyPlot.OxyColor.FromAColor(20, OxyPlot.OxyColors.LightYellow),
                Stroke = OxyPlot.OxyColors.Transparent,
            };
            polygonOuter.Points.AddRange(outerPoints);
            plotModel.Annotations.Add(polygonOuter);

            // --- 장애물 데이터가 있으면 극좌표 그래프에 표시 ---
            if (obstacleAzList != null && obstacleAltList != null && obstacleAzList.Count > 1 && obstacleAzList.Count == obstacleAltList.Count)
            {
                // 장애물 곡선
                var obsSeries = new OxyPlot.Series.LineSeries
                {
                    Color = OxyPlot.OxyColors.Orange,
                    StrokeThickness = 2,
                    Title = "지평선",
                    MarkerType = OxyPlot.MarkerType.Circle,
                    MarkerSize = 4,
                    MarkerFill = OxyPlot.OxyColors.OrangeRed
                };
                for (int i = 0; i < obstacleAzList.Count; i++)
                {
                    var point = new OxyPlot.DataPoint(90 - obstacleAltList[i], obstacleAzList[i]);
                    obsSeries.Points.Add(point);
                }
                plotModel.Series.Add(obsSeries);

                // 내부/외부 영역 폴리곤
                var areaPoints = new List<OxyPlot.DataPoint>();
                areaPoints.Add(new OxyPlot.DataPoint(90, 0));
                for (int i = 0; i < obstacleAzList.Count; i++)
                    areaPoints.Add(new OxyPlot.DataPoint(90 - obstacleAltList[i], obstacleAzList[i]));
                areaPoints.Add(new OxyPlot.DataPoint(90, 0));

                var polygonInner = new OxyPlot.Annotations.PolygonAnnotation
                {
                    Fill = OxyPlot.OxyColor.FromAColor(20, OxyPlot.OxyColors.LightYellow),
                    Stroke = OxyPlot.OxyColors.Transparent,
                };
                polygonInner.Points.AddRange(areaPoints);
                plotModel.Annotations.Add(polygonInner);

                var outerObsPoints = new List<OxyPlot.DataPoint>();
                for (int az = 0; az <= 360; az += 5)
                    outerObsPoints.Add(new OxyPlot.DataPoint(90, az));
                for (int i = areaPoints.Count - 1; i >= 0; i--)
                    outerObsPoints.Add(areaPoints[i]);
                var polygonOuterObs = new OxyPlot.Annotations.PolygonAnnotation
                {
                    Fill = OxyPlot.OxyColor.FromAColor(120, OxyPlot.OxyColors.DarkSlateBlue),
                    Stroke = OxyPlot.OxyColors.Transparent,
                };
                polygonOuterObs.Points.AddRange(outerObsPoints);
                plotModel.Annotations.Add(polygonOuterObs);
            }

            // 배경을 더 어둡게 설정
            plotModel.Background = OxyPlot.OxyColor.FromRgb(20, 20, 30);

            // ---- 가대 연결 시 현재 위치 노란색 점 표시 ----
            double? mountAz = null, mountAlt = null;
            try
            {
                if (telescope != null && telescope.Connected)
                {
                    mountAz = telescope.Azimuth;
                    mountAlt = telescope.Altitude;
                }
            }
            catch { /* 무시 */ }

            if (mountAz.HasValue && mountAlt.HasValue)
            {
                // 극좌표: (고도, 방위각)
                plotModel.Annotations.Add(new OxyPlot.Annotations.PointAnnotation
                {
                    X = 90 - mountAlt.Value, // OxyPlot 극좌표: 중심이 90, 가장자리가 0
                    Y = mountAz.Value,
                    Shape = OxyPlot.MarkerType.Circle,
                    Size = 10,
                    Fill = OxyPlot.OxyColors.Yellow,
                    Stroke = OxyPlot.OxyColors.Black,
                    StrokeThickness = 2,
                    ToolTip = "가대 현재 위치"
                });
            }

            // ---- 관측 순서 및 천체 위치 표시 ----
            if (finalObservationOrder != null && finalObservationOrder.Count > 0)
            {
                DateTime currentTime = GetSelectedDateTime();

                // 관측 순서 화살표 및 천체 위치 표시
                for (int i = 0; i < finalObservationOrder.Count; i++)
                {
                    var target = finalObservationOrder[i];

                    // 천체의 지평 좌표 계산
                    var equator = new AstroAlgo.Models.Equator
                    {
                        RA = target.RA * 15, // 시간 단위를 도 단위로 변환
                        Dec = target.Dec
                    };

                    double altitude = CalculateAltitude(target.RA * 15, target.Dec, site.lat, site.lon, currentTime);
                    double azimuth = CoordinateSystem.GetAzimuth(currentTime, equator, latitude: site.lat, longitude: site.lon);

                    // 고도가 0도 이상인 경우에만 표시
                    if (altitude > 0)
                    {
                        // 천체 위치 점 표시
                        plotModel.Annotations.Add(new OxyPlot.Annotations.PointAnnotation
                        {
                            X = 90 - altitude,
                            Y = azimuth,
                            Shape = OxyPlot.MarkerType.Circle,
                            Size = 4,
                            Fill = OxyPlot.OxyColors.Cyan,
                            Stroke = OxyPlot.OxyColors.White,
                            StrokeThickness = 1,
                            ToolTip = $"{target.Name}\n고도: {altitude:F1}°\n방위각: {azimuth:F1}°"
                        });

                        // 순서 번호 텍스트 표시
                        // plotModel.Annotations.Add(new OxyPlot.Annotations.TextAnnotation
                        // {
                        //    Text = (i + 1).ToString(),
                        //    TextPosition = new OxyPlot.DataPoint(90 - altitude, azimuth),
                        //    TextColor = OxyPlot.OxyColors.White,
                        //    FontSize = 10,
                        //    Background = OxyPlot.OxyColor.FromAColor(150, OxyPlot.OxyColors.Black),
                        //    TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Center,
                        //    TextVerticalAlignment = OxyPlot.VerticalAlignment.Middle,
                        //    Offset = new OxyPlot.ScreenVector(0, -15)
                        // });

                        // 화살표 표시 (다음 천체로의 방향)
                        if (i < finalObservationOrder.Count - 1)
                        {
                            var nextTarget = finalObservationOrder[i + 1];
                            var nextEquator = new AstroAlgo.Models.Equator
                            {
                                RA = nextTarget.RA * 15,
                                Dec = nextTarget.Dec
                            };

                            double nextAltitude = CalculateAltitude(nextTarget.RA * 15, nextTarget.Dec, site.lat, site.lon, currentTime);
                            double nextAzimuth = CoordinateSystem.GetAzimuth(currentTime, nextEquator, latitude: site.lat, longitude: site.lon);

                            if (nextAltitude > 0)
                            {
                                // 화살표 직선
                                plotModel.Annotations.Add(new OxyPlot.Annotations.ArrowAnnotation
                                {
                                    StartPoint = new OxyPlot.DataPoint(90 - altitude, azimuth),
                                    EndPoint = new OxyPlot.DataPoint(90 - nextAltitude, nextAzimuth),
                                    Color = OxyPlot.OxyColors.LightGreen,
                                    StrokeThickness = 1,
                                    HeadLength = 8,
                                    HeadWidth = 6
                                });
                            }
                        }
                    }
                }
            }

            // 렌더링
            var pngExporter = new PngExporter
            {
                Width = pictureBox4.Width,
                Height = pictureBox4.Height
            };
            using (var stream = new MemoryStream())
            {
                pngExporter.Export(plotModel, stream);
                stream.Position = 0;
                pictureBox4.Image?.Dispose();
                pictureBox4.Image = Image.FromStream(stream);
            }

            // === 2D 평면 그래프 생성 (방위각-고도) ==== //
            var plotModel2D = new OxyPlot.PlotModel
            {
                PlotAreaBorderColor = OxyPlot.OxyColors.Transparent,
                Background = OxyPlot.OxyColor.FromRgb(30, 30, 40),
                TextColor = OxyPlot.OxyColors.White,
            };

            // X축: 방위각 (0~360)
            plotModel2D.Axes.Add(new OxyPlot.Axes.LinearAxis
            {
                Position = OxyPlot.Axes.AxisPosition.Bottom,
                Minimum = 0,
                Maximum = 360,
                MajorStep = 45,
                MinorStep = 15,
                Title = "방위각 (°)",
                TextColor = OxyPlot.OxyColors.White,
                AxislineColor = OxyPlot.OxyColors.White,
                MajorGridlineStyle = OxyPlot.LineStyle.Solid,
                MajorGridlineColor = OxyPlot.OxyColor.FromRgb(80, 90, 100),
                FontSize = 11
            });

            // Y축: 고도 (0~90)
            plotModel2D.Axes.Add(new OxyPlot.Axes.LinearAxis
            {
                Position = OxyPlot.Axes.AxisPosition.Left,
                Minimum = 0,
                Maximum = 90,
                MajorStep = 30,
                MinorStep = 10,
                Title = "고도 (°)",
                TextColor = OxyPlot.OxyColors.White,
                AxislineColor = OxyPlot.OxyColors.White,
                MajorGridlineStyle = OxyPlot.LineStyle.Solid,
                MajorGridlineColor = OxyPlot.OxyColor.FromRgb(80, 90, 100),
                FontSize = 11
            });

            // -------------------- 자오선(남쪽, 방위각 180도) 굵은 선 추가 (2D) --------------------
            double lat2D = site.lat; // 서울 기본값, 필요시 변수로 대체
            var meridianLine2D = new OxyPlot.Series.LineSeries
            {
                Color = OxyPlot.OxyColors.Red,
                StrokeThickness = 2,
                LineStyle = OxyPlot.LineStyle.Solid,
                Title = "자오선"
            };
            // 수평선(고도 0, 방위각 180)
            meridianLine2D.Points.Add(new OxyPlot.DataPoint(180, 0));

            // 천정(고도 90, 방위각 180)
            meridianLine2D.Points.Add(new OxyPlot.DataPoint(180, 90));
            plotModel2D.Series.Add(meridianLine2D);

            // ---- 수정 예시: 두 개의 LineSeries로 각각 따로 그리기 ----
            var meridianLine2D_1 = new OxyPlot.Series.LineSeries
            {
                Color = OxyPlot.OxyColors.Red,
                StrokeThickness = 2,
                LineStyle = OxyPlot.LineStyle.Solid,
                Title = "자오선(수평선~천정)"
            };
            meridianLine2D_1.Points.Add(new OxyPlot.DataPoint(180, 0)); // 수평선(고도 0, 방위각 180)
            meridianLine2D_1.Points.Add(new OxyPlot.DataPoint(180, 90)); // 천정(고도 90, 방위각 180)
            plotModel2D.Series.Add(meridianLine2D_1);

            var meridianLine2D_2 = new OxyPlot.Series.LineSeries
            {
                Color = OxyPlot.OxyColors.Red,
                StrokeThickness = 2,
                LineStyle = OxyPlot.LineStyle.Solid,
                Title = "자오선(천정~북극)"
            };
            meridianLine2D_2.Points.Add(new OxyPlot.DataPoint(0, 90)); // 천정(고도 90, 방위각 0)
            meridianLine2D_2.Points.Add(new OxyPlot.DataPoint(0, lat2D)); // 천구의 북극(고도=위도, 방위각=0)
            plotModel2D.Series.Add(meridianLine2D_2);

            // 천구의 북극을 점으로 표시 (2D)
            var northPolePoint2D = new OxyPlot.Annotations.PointAnnotation
            {
                X = 0, // 방위각=0 (북쪽)
                Y = lat2D, // 고도=위도
                Shape = OxyPlot.MarkerType.Circle,
                Size = 8,
                Fill = OxyPlot.OxyColors.DeepSkyBlue,
                Stroke = OxyPlot.OxyColors.White,
                StrokeThickness = 2,
                ToolTip = "천구의 북극"
            };
            plotModel2D.Annotations.Add(northPolePoint2D);

            // --- 장애물 데이터가 있으면 2D 그래프에도 표시 ---
            if (obstacleAzList != null && obstacleAltList != null && obstacleAzList.Count > 1 && obstacleAzList.Count == obstacleAltList.Count)
            {
                var obsLine2D = new OxyPlot.Series.LineSeries
                {
                    Color = OxyPlot.OxyColors.Orange,
                    StrokeThickness = 2,
                    Title = "지평선",
                    MarkerType = OxyPlot.MarkerType.Circle,
                    MarkerSize = 4,
                    MarkerFill = OxyPlot.OxyColors.OrangeRed
                };
                for (int i = 0; i < obstacleAzList.Count; i++)
                {
                    obsLine2D.Points.Add(new OxyPlot.DataPoint(obstacleAzList[i], obstacleAltList[i]));
                }
                plotModel2D.Series.Add(obsLine2D);

                var areaSeries = new OxyPlot.Series.AreaSeries
                {
                    Color = OxyPlot.OxyColor.FromAColor(120, OxyPlot.OxyColors.DarkSlateBlue),
                    MarkerType = OxyPlot.MarkerType.None
                };
                for (int i = 0; i < obstacleAzList.Count; i++)
                {
                    areaSeries.Points.Add(new OxyPlot.DataPoint(obstacleAzList[i], obstacleAltList[i]));
                    areaSeries.Points2.Add(new OxyPlot.DataPoint(obstacleAzList[i], 0));
                }
                plotModel2D.Series.Add(areaSeries);
            }

            // ---- 가대 연결 시 현재 위치 노란색 점 표시 (2D) ----
            if (mountAz.HasValue && mountAlt.HasValue)
            {
                plotModel2D.Annotations.Add(new OxyPlot.Annotations.PointAnnotation
                {
                    X = mountAz.Value,
                    Y = mountAlt.Value,
                    Shape = OxyPlot.MarkerType.Circle,
                    Size = 10,
                    Fill = OxyPlot.OxyColors.Yellow,
                    Stroke = OxyPlot.OxyColors.Black,
                    StrokeThickness = 2,
                    ToolTip = "가대 현재 위치"
                });
            }

            // ---- 관측 순서 및 천체 위치 표시 (2D) ----
            if (finalObservationOrder != null && finalObservationOrder.Count > 0)
            {
                DateTime currentTime = GetSelectedDateTime();

                // 관측 순서 화살표 및 천체 위치 표시
                for (int i = 0; i < finalObservationOrder.Count; i++)
                {
                    var target = finalObservationOrder[i];

                    // 천체의 지평 좌표 계산
                    var equator = new AstroAlgo.Models.Equator
                    {
                        RA = target.RA * 15, // 시간 단위를 도 단위로 변환
                        Dec = target.Dec
                    };

                    double altitude = CalculateAltitude(target.RA * 15, target.Dec, site.lat, site.lon, currentTime);
                    double azimuth = CoordinateSystem.GetAzimuth(currentTime, equator, latitude: site.lat, longitude: site.lon);
                    // Console.WriteLine($"Target: {target.Name}, RA: {target.RA}, Dec: {target.Dec}, Altitude: {altitude}, Azimuth: {azimuth}");
                    // 고도가 0도 이상인 경우에만 표시
                    if (altitude > 0)
                    {
                        // 천체 위치 점 표시
                        plotModel2D.Annotations.Add(new OxyPlot.Annotations.PointAnnotation
                        {
                            X = azimuth,
                            Y = altitude,
                            Shape = OxyPlot.MarkerType.Circle,
                            Size = 4,
                            Fill = OxyPlot.OxyColors.Cyan,
                            Stroke = OxyPlot.OxyColors.White,
                            StrokeThickness = 1,
                            ToolTip = $"{target.Name}\n고도: {altitude:F1}°\n방위각: {azimuth:F1}°"
                        });

                        // 순서 번호 텍스트 표시
                        //
                        // plotModel2D.Annotations.Add(new OxyPlot.Annotations.TextAnnotation
                        // {
                        //    Text = (i + 1).ToString(),
                        //    TextPosition = new OxyPlot.DataPoint(azimuth, altitude),
                        //    TextColor = OxyPlot.OxyColors.White,
                        //    FontSize = 10,
                        //    Background = OxyPlot.OxyColor.FromAColor(150, OxyPlot.OxyColors.Black),
                        //    TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Center,
                        //    TextVerticalAlignment = OxyPlot.VerticalAlignment.Middle,
                        //    Offset = new OxyPlot.ScreenVector(0, -15)
                        // });
                        // 

                        // 화살표 표시 (다음 천체로의 방향)
                        if (i < finalObservationOrder.Count - 1)
                        {
                            var nextTarget = finalObservationOrder[i + 1];
                            var nextEquator = new AstroAlgo.Models.Equator
                            {
                                RA = nextTarget.RA * 15,
                                Dec = nextTarget.Dec
                            };

                            double nextAltitude = CalculateAltitude(nextTarget.RA * 15, nextTarget.Dec, site.lat, site.lon, currentTime);
                            double nextAzimuth = CoordinateSystem.GetAzimuth(currentTime, nextEquator, latitude: site.lat, longitude: site.lon);

                            if (nextAltitude > 0)
                            {
                                // 화살표 직선
                                plotModel2D.Annotations.Add(new OxyPlot.Annotations.ArrowAnnotation
                                {
                                    StartPoint = new OxyPlot.DataPoint(azimuth, altitude),
                                    EndPoint = new OxyPlot.DataPoint(nextAzimuth, nextAltitude),
                                    Color = OxyPlot.OxyColors.LightGreen,
                                    StrokeThickness = 1,
                                    HeadLength = 8,
                                    HeadWidth = 6
                                });
                            }
                        }
                    }
                }
            }

            // 렌더링
            var pngExporter2D = new OxyPlot.WindowsForms.PngExporter
            {
                Width = pictureBox5.Width,
                Height = pictureBox5.Height
            };
            using (var stream2D = new MemoryStream())
            {
                pngExporter2D.Export(plotModel2D, stream2D);
                stream2D.Position = 0;
                pictureBox5.Image?.Dispose();
                pictureBox5.Image = Image.FromStream(stream2D);
            }
        }


        //============================================================================//


        // 천체의 시간 - 고도 그래프 그리기 (KST 기준)
        private void GenerateAltitudePlot(double ra, double dec, double siteLat, double siteLon, DateTime date, string objectName)
        {
            try
            {
                // 1. KST 기준 시간축 생성 (12:00~다음날 12:00, 3분 간격) -> 요놈은 천체 고도 및 태양 고도 계산에 사용됨. 구간 정해서 간격마다의 고도 구하고 피팅하여 곡선 그림
                var times = new List<DateTime>();
                var altitudes = new List<double>();
                DateTime kstStart = date.Date.AddHours(12); // KST 12:00
                DateTime kstEnd = kstStart.AddHours(24);

                for (DateTime kstTime = kstStart; kstTime <= kstEnd; kstTime = kstTime.AddMinutes(3))
                {
                    times.Add(kstTime);
                }


                // 2. 천체 고도 계산
                foreach (var kstTime in times)
                {
                    DateTime utcTime = kstTime.AddHours(0); // KST
                    double raDegree = ra; // RA는 도 단위로 입력됨
                    double altitude = CalculateAltitude(raDegree, dec, siteLat, siteLon, utcTime); // 얘는 적경, 적위 모두 '도' 단위이어야함
                    altitudes.Add(Math.Max(0, altitude)); // 음수값은 0으로 처리
                }


                // 3. 태양 고도 계산
                var sunAltitudes = new List<double>();
                foreach (var kstTime in times)
                {
                    DateTime utcTime = kstTime.AddHours(0); // KST
                    double sunAlt = CalculateSunAltitude(siteLat, siteLon, utcTime);
                    sunAltitudes.Add(sunAlt);
                }


                // 4. OxyPlot 그래프 생성
                var plotModel = new PlotModel
                {
                    PlotAreaBorderColor = OxyColors.Transparent,
                    Background = OxyColor.FromRgb(45, 55, 65),
                    TextColor = OxyColors.White
                };


                // X축: KST 시간
                var xAxis = new LinearAxis
                {
                    Position = AxisPosition.Bottom,
                    Minimum = 12,
                    Maximum = 36,
                    MajorStep = 3,
                    MinorStep = 1,
                    TextColor = OxyColors.White,
                    AxislineColor = OxyColors.Gray,
                    MajorGridlineStyle = LineStyle.Solid,
                    MajorGridlineColor = OxyColor.FromRgb(80, 90, 100),
                    FontSize = 11,
                    LabelFormatter = value =>
                    {
                        int hour = ((int)value) % 24;
                        return hour.ToString("00");
                    }
                };
                plotModel.Axes.Add(xAxis);


                // Y축: 고도
                var yAxis = new LinearAxis
                {
                    Position = AxisPosition.Left,
                    Minimum = 0,
                    Maximum = 90,
                    MajorStep = 30,
                    MinorStep = 10,
                    TextColor = OxyColors.White,
                    AxislineColor = OxyColors.Gray,
                    MajorGridlineStyle = LineStyle.Solid,
                    MajorGridlineColor = OxyColor.FromRgb(80, 90, 100),
                    FontSize = 11
                };
                plotModel.Axes.Add(yAxis);


                // 5. 박명 영역 추가
                AddAccurateTwilightRegions(plotModel, times, sunAltitudes);


                // 6. 천체 고도 곡선 표시
                var series = new LineSeries
                {
                    Color = OxyColors.White,
                    StrokeThickness = 2
                };

                for (int i = 0; i < times.Count; i++)
                {
                    double hour = (times[i] - kstStart.Date).TotalHours;
                    series.Points.Add(new DataPoint(hour, altitudes[i]));
                }
                plotModel.Series.Add(series);


                // 7. 현재 KST 시간 표시 (점선)
                DateTime nowKst = DateTime.Now; // 시스템이 KST라고 가정
                double nowHour = (nowKst - kstStart.Date).TotalHours;
                if (nowHour >= 12 && nowHour <= 36)
                {
                    plotModel.Annotations.Add(new LineAnnotation
                    {
                        Type = LineAnnotationType.Vertical,
                        X = nowHour,
                        Color = OxyColors.Red,
                        LineStyle = LineStyle.Dash,
                        StrokeThickness = 2
                    });

                    // 현재 시각 텍스트 표시
                    plotModel.Annotations.Add(new TextAnnotation
                    {
                        Text = $"{nowKst:HH:mm}",
                        TextPosition = new DataPoint(nowHour, 40),
                        TextColor = OxyColors.Red,
                        FontSize = 10,
                        Background = OxyColor.FromAColor(120, OxyColors.Black)
                    });
                }


                // 8. 최대고도 표시
                int maxIdx = altitudes.IndexOf(altitudes.Max());
                if (maxIdx >= 0)
                {
                    double maxHour = (times[maxIdx] - kstStart.Date).TotalHours;
                    double maxAlt = altitudes[maxIdx];

                    plotModel.Annotations.Add(new PointAnnotation
                    {
                        X = maxHour,
                        Y = maxAlt,
                        Shape = MarkerType.Circle,
                        Size = 6,
                        Fill = OxyColors.White
                    });

                    if (maxAlt < 35)
                    {
                        hight = maxAlt + 24;
                    }
                    else
                    {
                        hight = maxAlt - 16;
                    }

                    // 텍스트를 최대고도 점 바로 아래에 배치
                    plotModel.Annotations.Add(new TextAnnotation
                    {
                        Text = $"{Math.Round(maxAlt, 1)}°",
                        TextPosition = new DataPoint(maxHour, Math.Max(hight, 0)), // 점 바로 아래, 0도 미만 방지
                        TextColor = OxyColors.White,
                        FontSize = 10,
                        Background = OxyColor.FromAColor(100, OxyColors.Black),
                        TextVerticalAlignment = VerticalAlignment.Top
                    });
                }

                // 9. pictureBox3에 렌더링
                var pngExporter = new PngExporter
                {
                    Width = pictureBox3.Width,
                    Height = pictureBox3.Height
                };

                using (var stream = new MemoryStream())
                {
                    pngExporter.Export(plotModel, stream);
                    stream.Position = 0;

                    pictureBox3.Image?.Dispose();
                    pictureBox3.Image = Image.FromStream(stream);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"그래프 생성 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

       
        // 태양 고도 계산 메서드 -> 박명 영역 계산용
        private double CalculateSunAltitude(double lat, double lon, DateTime utc)
        {
            // Astronomy 라이브러리의 SunPosition 객체 사용
            var sunPos = SunCalculator.GetSunPosition(utc, lat, lon);

            // 라이브러리를 통해 구한 고도는 radians 단위이므로, degrees로 변환
            double altitudeDeg = sunPos.Altitude * 180.0 / Math.PI;

            return altitudeDeg;
        }


        // 천체 고도 계산 메서드 (utc라 써 놓긴 했는데 실제로는 kst로 사용됨)
        private double CalculateAltitude(double ra, double dec, double lat, double lon, DateTime utc)
        {
            // AstroAlgo의 Equator 객체 생성 (RA: 도 단위, Dec: 도 단위)
            var equator = new Equator
            {
                RA = ra,   // 반드시 도 단위로 변환해서 입력 (예: 10h → 150도)
                Dec = dec  // 도 단위
            };

            // 고도(고도각) 계산
            double altitude = CoordinateSystem.GetElevationAngle(utc, equator, latitude: lat, longitude: lon);
            return altitude;
        }


        // 박명 영역 추가 메서드
        private void AddAccurateTwilightRegions(PlotModel plotModel, List<DateTime> times, List<double> sunAltitudes)
        {
            for (int i = 0; i < times.Count - 1; i++)
            {
                double hour1 = (times[i] - times[0].Date).TotalHours;
                double hour2 = (times[i + 1] - times[0].Date).TotalHours;
                double sunAlt = sunAltitudes[i];

                OxyColor regionColor;

                if (sunAlt < -18) // 천문박명 (완전한 밤)
                {
                    regionColor = OxyColor.FromAColor(150, OxyColors.Black);
                }
                else if (sunAlt < -12) // 항해박명
                {
                    regionColor = OxyColor.FromAColor(100, OxyColors.DarkBlue); 
                }
                else if (sunAlt < -6) // 상용박명 (시민박명)
                {
                    regionColor = OxyColor.FromAColor(80, OxyColors.Orange);
                }
                else if (sunAlt < 0) // 일출/일몰 직전후
                {
                    regionColor = OxyColor.FromAColor(60, OxyColors.Orange);
                }
                else // 낮
                {
                    regionColor = OxyColor.FromAColor(40, OxyColors.LightBlue);
                }

                var rectangle = new RectangleAnnotation
                {
                    MinimumX = hour1,
                    MaximumX = hour2,
                    MinimumY = 0,
                    MaximumY = 90,
                    Fill = regionColor,
                    Stroke = OxyColors.Transparent
                };
                plotModel.Annotations.Add(rectangle);
            }
        }


        //============================================================================//


        // ASCOM 드라이버 연결 버튼
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                // 가대를 선택하는 창을 띄운 후, 선택된 가대의 ID를 받아온다.
                string id = Telescope.Choose("");

                // 가대를 선택하지 않은 경우, 중단한다.
                if (string.IsNullOrEmpty(id))
                {
                    UpdateTelescopeStatus(); // 상태 업데이트
                    return;
                }

                // 기존 연결 해제
                DisconnectTelescope();

                // 선택된 가대의 객체를 생성한다.
                telescope = new Telescope(id);
                textBox2.Text = id;

                finder.SetTelescope(telescope);

                // 가대에 연결한다.
                telescope.Connected = true;

                // 가대 연결 상태 업데이트
                UpdateTelescopeStatus();

                // 관측지 정보를 한 번만 가져옴
                GetTelescopeLocationData();

                // Sky180Viewer 타이머 설정
                if (GenerateSkyPlotTimer == null)
                {
                    GenerateSkyPlotTimer = new Timer();
                    GenerateSkyPlotTimer.Interval = 1000; // 1초
                    GenerateSkyPlotTimer.Tick += (s, ev) => GenerateSkyPlot();
                }

                // GenerateHorizonPlotTimer가 작동 중이면 GenerateSkyPlotTimer를 시작하지 않음
                if (GenerateHorizonPlotTimer == null || !GenerateHorizonPlotTimer.Enabled)
                {
                    GenerateSkyPlotTimer.Start();
                }

                // tabPage6가 열리도록 설정
                if (tabControl2 != null && tabPage6 != null)
                {
                    tabControl2.SelectedTab = tabPage6;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"가대 연결 중 오류가 발생했습니다: {ex.Message}", "연결 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateTelescopeStatus(); // 오류 발생 시에도 상태 업데이트

                // 오류 발생시 연결 해제
                DisconnectTelescope();

                // 타이머 중지
                if (GenerateSkyPlotTimer != null)
                    GenerateSkyPlotTimer.Stop();
            }
        }


        // 가대 연결 해제 메서드 DisconnectTelescope에서 타이머도 중지
        private void DisconnectTelescope()
        {
            try
            {
                // ASCOM 연결 해제
                if (telescope != null)
                {
                    if (telescope.Connected)
                    {
                        telescope.Connected = false;
                    }
                    telescope.Dispose();
                    telescope = null; // 참조 제거
                }

                // 연결 해제 시 텍스트박스 초기화
                DisplayInitialValues();

                finder.SetTelescope(null);

                // Sky180Viewer 타이머 중지
                if (GenerateSkyPlotTimer != null)
                    GenerateSkyPlotTimer.Stop();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"가대 연결 해제 오류: {ex.Message}");
            }
        }


        // 가대 연결 시 관측지 정보 가져오기
        private void GetTelescopeLocationData()
        {
            try
            {
                if (telescope != null && telescope.Connected)
                {
                    // ASCOM에서 관측지 정보를 한 번만 가져옴
                    double latitude = telescope.SiteLatitude;
                    double longitude = telescope.SiteLongitude;
                    double altitude = telescope.SiteElevation;

                    textBox3.ReadOnly = true;
                    textBox4.ReadOnly = true;
                    textBox6.ReadOnly = true;

                    button25.Enabled = false;

                    label3.Text = "관측지 위도(도)";
                    label4.Text = "관측지 경도(도)";
                    label12.Text = "관측지 고도(m)";

                    // UI에 표시
                    textBox3.Text = $"{latitude:F6}";
                    textBox4.Text = $"{longitude:F6}";
                    textBox6.Text = $"{altitude:F2}";
                }
            }
            catch (Exception ex)
            {
                // 오류 발생 시 오류 메시지 표시
                textBox3.Text = "정보 없음";
                textBox4.Text = "정보 없음";
                textBox6.Text = "정보 없음";
                Console.WriteLine($"관측지 정보 가져오기 오류: {ex.Message}");
            }
        }


        // 가대 연결 상태 업데이트 메서드
        private void UpdateTelescopeStatus()
        {
            try
            {
                if (telescope != null && telescope.Connected)
                {
                    label2.Text = "연결됨";
                    label2.ForeColor = Color.Green;
                }
                else
                {
                    label2.Text = "연결 안됨";
                    label2.ForeColor = Color.Red;
                }
            }
            catch (Exception)
            {
                label2.Text = "상태 불명";
                label2.ForeColor = Color.Orange;
            }
        }


        // 가대 연결 해제 버튼
        private void button26_Click(object sender, EventArgs e)
        {
            DisconnectTelescope();
            UpdateTelescopeStatus();
            DisplayInitialValues();

            if (GenerateHorizonPlotTimer != null)
                GenerateHorizonPlotTimer.Stop();
        }


        // 위도 / 경도 / 고도 수동 입력 버튼
        private void button25_Click(object sender, EventArgs e)
        {
            GenerateSkyPlot();
        }


        //============================================================================//


        // 검색 결과 UI 초기화 -> 처음에 실행하는 용도
        private void InitializeSearchResults()
        {
            // 검색 결과 바인딩 시
            listBox1.DisplayMember = "Names";


            // 이벤트 핸들러 중복 등록 방지
            if (!isEventHandlerRegistered)
            {
                listBox1.SelectedIndexChanged += ListBox1_SelectedIndexChanged;
                listBox1.SelectedIndexChanged += StarListBox1_SelectedIndexChanged;
                isEventHandlerRegistered = true;
            }

            // 디자이너에서 생성한 라벨들 초기화
            ClearDetailPanel();
        }


        // 검색 정보 패널 초기화 -> 처음에 실행하는 용도
        private void ClearDetailPanel()
        {
            label5.Text = "천체명:";
            label6.Text = "종류:";
            label7.Text = "적경:";
            label8.Text = "적위:";
            label9.Text = "겉보기 등급:";
            label10.Text = "별자리:";
        }


        // 검색 시작 버튼
        private void button2_Click(object sender, EventArgs e)
        {


            string searchTerm = textBox1.Text.Trim();
            if (string.IsNullOrEmpty(searchTerm))
            {
                MessageBox.Show("검색어를 입력해주세요.", "검색", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                if (radioButton1.Checked) // 딥스카이 검색 모드
                {
                    List<DeepSkyObject> searchResults = SearchNGCCatalog(searchTerm);

                    listBox1.BeginUpdate();
                    listBox1.DataSource = null;
                    listBox1.DataSource = searchResults;
                    listBox1.DisplayMember = "Name";
                    listBox1.EndUpdate();

                    // 이벤트 핸들러 설정
                    listBox1.SelectedIndexChanged -= StarListBox1_SelectedIndexChanged;
                    listBox1.SelectedIndexChanged -= ListBox1_SelectedIndexChanged;
                    listBox1.SelectedIndexChanged += ListBox1_SelectedIndexChanged;

                    if (searchResults.Count == 0)
                    {
                        MessageBox.Show("검색 결과가 없습니다.", "검색", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        ClearDetailPanel();
                    }
                    else
                    {
                        listBox1.SelectedIndex = 0;
                    }

                    // tabPage5가 열리도록 설정
                    if (tabControl2 != null && tabPage5 != null)
                    {
                        tabControl2.SelectedTab = tabPage5;
                    }
                }
                else if (radioButton2.Checked) // 별 검색 모드
                {
                    var starResults = SearchStarCatalog(searchTerm);

                    listBox1.BeginUpdate();
                    listBox1.DataSource = null;
                    listBox1.DataSource = starResults;
                    listBox1.DisplayMember = "Names";
                    listBox1.EndUpdate();

                    // 이벤트 핸들러 설정
                    listBox1.SelectedIndexChanged -= ListBox1_SelectedIndexChanged;
                    listBox1.SelectedIndexChanged -= StarListBox1_SelectedIndexChanged;
                    listBox1.SelectedIndexChanged += StarListBox1_SelectedIndexChanged;

                    if (starResults.Count == 0)
                    {
                        MessageBox.Show("검색 결과가 없습니다.", "검색", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        ClearDetailPanel();
                    }
                    else
                    {
                        listBox1.SelectedIndex = 0;
                    }

                    // tabPage5가 열리도록 설정
                    if (tabControl2 != null && tabPage5 != null)
                    {
                        tabControl2.SelectedTab = tabPage5;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"검색 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // 별 카탈로그 검색
        private List<StarObject> SearchStarCatalog(string searchTerm)
        {
            var results = new List<StarObject>();
            var seenIds = new HashSet<string>();
            string csvFilePath = Path.Combine(Application.StartupPath, "hyg_v42.csv");

            if (!File.Exists(csvFilePath))
            {
                MessageBox.Show("hyg_v42.csv 파일을 찾을 수 없습니다.", "파일 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return results;
            }

            try
            {
                using (var reader = new StreamReader(csvFilePath, Encoding.UTF8))
                {
                    string line;
                    bool firstLine = true;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (firstLine)
                        {
                            firstLine = false; // 헤더 스킵
                            continue;
                        }
                        if (string.IsNullOrWhiteSpace(line)) continue;

                        string[] fields = line.Split(',');

                        if (fields.Length < 14) continue;

                        string id = fields[0].Trim('"');
                        if (seenIds.Contains(id)) continue; // 중복 방지

                        string proper = fields[6].Trim('"');
                        string bf = fields[5].Trim('"');
                        string hd = fields[2].Trim('"');
                        string hip = fields[1].Trim('"');

                        // 이름 목록 생성
                        var nameList = new List<string>();
                        if (!string.IsNullOrEmpty(proper)) nameList.Add(proper);
                        if (!string.IsNullOrEmpty(bf)) nameList.Add(bf);
                        if (!string.IsNullOrEmpty(hd)) nameList.Add("HD " + hd);
                        if (!string.IsNullOrEmpty(hip)) nameList.Add("HIP " + hip);

                        // 검색어가 이름 중 하나에 포함되면 결과에 추가
                        if (nameList.Any(n => n.ToLower().Contains(searchTerm.ToLower())))
                        {
                            string ra = fields[7];
                            string dec = fields[8];
                            string magStr = fields[13].Trim();
                            double mag = 0;
                            bool magValid = double.TryParse(magStr, out mag);
                            string con = fields[29].Trim(); // con 열(별자리) 읽기

                            results.Add(new StarObject
                            {
                                Names = string.Join(" / ", nameList),
                                RA = ra,
                                Dec = dec,
                                Magnitude = (magValid ? mag : double.NaN), // 파싱 실패(예: '-')면 NaN
                                Constellation = con // 추가
                            });
                            seenIds.Add(id);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"별 데이터 파일 읽기 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return results.OrderBy(x => x.Names).ToList();
        }


        // 별 검색 결과 라벨에 띄우기
        private void StarListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem is StarObject selectedStar)
            {
                label5.Text = $"이름: {selectedStar.Names}";
                label7.Text = $"적경: {selectedStar.RA}";
                label8.Text = $"적위: {selectedStar.Dec}";
                label9.Text = $"겉보기 등급: {(!double.IsNaN(selectedStar.Magnitude) ? selectedStar.Magnitude.ToString("F2") : "정보 없음")}";
                label10.Text = $"별자리: {selectedStar.Constellation}";
                label6.Text = "";

                if (double.TryParse(selectedStar.RA, out double ra) &&
                    double.TryParse(selectedStar.Dec, out double dec))
                {
                    var site = GetObservingSite();
                    GenerateAltitudePlot(ra * 15, dec, site.lat, site.lon, DateTime.Now, selectedStar.Names);
                }

                // tabPage5가 열리도록 설정
                if (tabControl2 != null && tabPage5 != null)
                {
                    tabControl2.SelectedTab = tabPage5;
                }
            }
        }


        // 딥스카이 카탈로그 검색
        private List<DeepSkyObject> SearchNGCCatalog(string searchTerm)
        {
            List<DeepSkyObject> results = new List<DeepSkyObject>();

            string csvFilePath = Path.Combine(Application.StartupPath, "NGC.csv");

            if (!File.Exists(csvFilePath))
            {
                MessageBox.Show("NGC.csv 파일을 찾을 수 없습니다.", "파일 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return results;
            }

            try
            {
                // StreamReader를 사용하여 메모리 사용량 최적화
                using (var reader = new StreamReader(csvFilePath, Encoding.UTF8))
                {
                    string line;
                    bool firstLine = true;

                    while ((line = reader.ReadLine()) != null)
                    {
                        if (firstLine)
                        {
                            firstLine = false; // 헤더 스킵
                            continue;
                        }

                        if (string.IsNullOrWhiteSpace(line)) continue;

                        string[] fields = line.Split(';');
                        if (fields.Length < 10) continue;


                        string name = fields[0].Trim();
                        string Messier = fields[23].TrimStart('0').Trim();

                        var nameList = new List<string>();
                        if (!string.IsNullOrEmpty(name)) nameList.Add(name);
                        if (!string.IsNullOrEmpty(Messier)) nameList.Add("M" + Messier);

                        if (nameList.Any(n => n.ToLower().Contains(searchTerm.ToLower())))
                        {
                            DeepSkyObject obj = new DeepSkyObject
                            {
                                Name = string.Join(" / ", nameList),
                                // 기존 코드
                                // Type = fields[1].Trim(),

                                // 수정된 코드 (한글 매핑)
                                Type = MapObjectTypeToKorean(fields[1].Trim()),

                                // 아래에 메서드 추가 (클래스 내부 아무 곳에나)

                                RA = fields[2].Trim(),
                                Dec = fields[3].Trim(),
                                Magnitude = double.TryParse(fields[8], out double mag) ? mag : 0,
                                Constellation = fields[4].Trim()
                            };
                            results.Add(obj);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"파일 읽기 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return results.OrderBy(x => x.Name).ToList();
        }


        // 딥스카이 검색 결과 라벨에 띄우기
        private void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem is DeepSkyObject selectedObject)
            {

                // NGC 데이터는 시분초 형식이므로 전용 파싱 함수 사용
                double ra = ParseRA(selectedObject.RA);    // 시분초 → 시간
                double dec = ParseDec(selectedObject.Dec); // 도분초 → 도

                label5.Text = $"천체명: {selectedObject.Name}";
                label6.Text = $"종류: {selectedObject.Type}";
                label7.Text = $"적경: {ra:F6}";
                label8.Text = $"적위: {dec:F6}";
                label9.Text = $"겉보기 등급: {(selectedObject.Magnitude > 0 ? selectedObject.Magnitude.ToString("F1") : "정보 없음")}";
                label10.Text = $"별자리: {selectedObject.Constellation}";

                if (ra != 0 || dec != 0) // 파싱이 성공했다면
                {
                    var site = GetObservingSite();
                    GenerateAltitudePlot(ra * 15, dec, site.lat, site.lon, DateTime.Now, selectedObject.Name);
                }

                // tabPage5가 열리도록 설정
                if (tabControl2 != null && tabPage5 != null)
                {
                    tabControl2.SelectedTab = tabPage5;
                }
            }
        }


        // 천체 종류를 한글로 매핑하는 메서드
        private string MapObjectTypeToKorean(string type)
        {
            switch (type)
            {
                case "*": return "별";
                case "**": return "쌍성";
                case "*Ass": return "성협";
                case "OCl": return "산개성단";
                case "GCl": return "구상성단";
                case "Cl+N": return "성단+성운";
                case "G": return "은하";
                case "GPair": return "이중 은하";
                case "GTrpl": return "삼중 은하";
                case "GGroup": return "은하군";
                case "PN": return "행성상성운";
                case "HII": return "HII 이온화 영역";
                case "DrkN": return "암흑성운";
                case "EmN": return "방출성운";
                case "Neb": return "성운";
                case "RfN": return "반사성운";
                case "SNR": return "초신성잔해";
                case "Nova": return "신성";
                case "NonEx": return "존재하지 않음";
                case "Dup": return "중복 객체";
                case "Other": return "기타";
                default: return type; // 알 수 없는 경우 원본 반환
            }
        }


        // 적경을 시분초에서 시간(소수점)으로 변환 -> 딥스카이 카탈로그 검색에서만 사용
        private double ParseRA(string raString)
        {
            try
            {
                // "00:00:00.00" 형식을 파싱
                var parts = raString.Split(':');
                if (parts.Length >= 2)
                {
                    double hours = double.Parse(parts[0]);
                    double minutes = double.Parse(parts[1]);
                    double seconds = parts.Length > 2 ? double.Parse(parts[2]) : 0;

                    return hours + minutes / 60.0 + seconds / 3600.0;
                }
            }
            catch
            {
                // 파싱 실패시 0 반환
            }
            return 0;
        }


        // 적위를 도분초에서 도(소수점)로 변환 -> 딥스카이 카탈로그 검색에서만 사용
        private double ParseDec(string decString)
        {
            try
            {
                // "+00:00:00.0" 또는 "-00:00:00.0" 형식을 파싱
                bool isNegative = decString.StartsWith("-");
                string cleanDec = decString.TrimStart('+', '-');

                var parts = cleanDec.Split(':');
                if (parts.Length >= 2)
                {
                    double degrees = double.Parse(parts[0]);
                    double minutes = double.Parse(parts[1]);
                    double seconds = parts.Length > 2 ? double.Parse(parts[2]) : 0;

                    double result = degrees + minutes / 60.0 + seconds / 3600.0;
                    return isNegative ? -result : result;
                }
            }
            catch
            {
                // 파싱 실패시 0 반환
            }
            return 0;
        }


        // 관측대상 목록에 추가 버튼 (관측 시간은 listBox2에서 직접 입력)
        private void button3_Click(object sender, EventArgs e)
        {
            string name = null;
            double ra = 0, dec = 0;

            if (radioButton1.Checked && listBox1.SelectedItem is DeepSkyObject dso)
            {
                name = dso.Name;
                ra = ParseRA(dso.RA); // 시→도
                dec = ParseDec(dso.Dec);
            }
            else if (radioButton2.Checked && listBox1.SelectedItem is StarObject star)
            {
                name = star.Names;
                if (double.TryParse(star.RA, out double raVal) && double.TryParse(star.Dec, out double decVal))
                {
                    ra = raVal; // 시→도
                    dec = decVal;
                }
                else
                {
                    MessageBox.Show("좌표 정보가 올바르지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            else
            {
                MessageBox.Show("대상을 선택하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 중복 방지: 같은 이름+좌표가 이미 있으면 추가하지 않음
            if (observationTargets.Any(t => t.Name == name && Math.Abs(t.RA - ra) < 1e-6 && Math.Abs(t.Dec - dec) < 1e-6))
            {
                MessageBox.Show("이미 목록에 추가된 대상입니다.", "중복", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 관측 시간은 기본값 10분으로 추가, listBox2에서 직접 수정 가능
            double obsTime = 10; // 기본값
            if (double.TryParse(textBox5.Text, out double userObsTime) && userObsTime > 0)
            {
                obsTime = userObsTime;
            }
            observationTargets.Add(new ObservationTarget
            {
                Name = name,
                RA = ra,
                Dec = dec,
                ObsTime = obsTime // textBox5 값 사용
            });

            // listBox2에 표시 (DataSource 재설정)
            listBox2.DataSource = null;
            listBox2.DataSource = observationTargets;
            listBox2.DisplayMember = "ToString";
        }


        // 관측 대상 목록에서 선택된 항목 삭제 버튼
        private void button9_Click(object sender, EventArgs e)
        {
            // listBox2에서 선택된 항목 삭제
            int selectedIndex = listBox2.SelectedIndex;
            if (selectedIndex >= 0 && selectedIndex < observationTargets.Count)
            {
                observationTargets.RemoveAt(selectedIndex);
                listBox2.DataSource = null;
                listBox2.DataSource = observationTargets;
                listBox2.DisplayMember = "ToString";
            }
        }


        // 관측 대상 목록 저장 버튼
        private void button10_Click(object sender, EventArgs e)
        {
            if (observationTargets.Count == 0)
            {
                MessageBox.Show("저장할 관측 대상이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "관측 대상 목록 (*.txt)|*.txt|모든 파일 (*.*)|*.*";
                saveFileDialog.Title = "관측 대상 목록 저장";
                saveFileDialog.FileName = "observation_targets.txt";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (var writer = new StreamWriter(saveFileDialog.FileName, false, Encoding.UTF8))
                        {
                            // 헤더
                            writer.WriteLine("#Name|RA|Dec|ObsTime(min)");
                            foreach (var target in observationTargets)
                            {
                                // 구분자: | (파이프)
                                writer.WriteLine($"{target.Name}|{target.RA}|{target.Dec}|{target.ObsTime}");
                            }
                        }
                        MessageBox.Show("관측 대상 목록이 저장되었습니다.", "저장 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"저장 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }


        // 관측 대상 목록 불러오기 버튼
        private void button11_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "관측 대상 목록 (*.txt)|*.txt|모든 파일 (*.*)|*.*";
                openFileDialog.Title = "관측 대상 목록 불러오기";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var loadedTargets = new List<ObservationTarget>();
                        using (var reader = new StreamReader(openFileDialog.FileName, Encoding.UTF8))
                        {
                            string line;
                            while ((line = reader.ReadLine()) != null)
                            {
                                if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#")) continue;
                                var parts = line.Split('|');
                                if (parts.Length < 4) continue;

                                string name = parts[0];
                                if (!double.TryParse(parts[1], out double ra)) continue;
                                if (!double.TryParse(parts[2], out double dec)) continue;
                                if (!double.TryParse(parts[3], out double obsTime)) obsTime = 10;

                                loadedTargets.Add(new ObservationTarget
                                {
                                    Name = name,
                                    RA = ra,
                                    Dec = dec,
                                    ObsTime = obsTime
                                });
                            }
                        }
                        observationTargets.Clear();
                        observationTargets.AddRange(loadedTargets);
                        listBox2.DataSource = null;
                        listBox2.DataSource = observationTargets;
                        listBox2.DisplayMember = "ToString";
                        MessageBox.Show("관측 대상 목록을 불러왔습니다.", "불러오기 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"불러오기 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }


        // 수동 관측 대상 추가 버튼
        private void button34_Click(object sender, EventArgs e)
        {
            using (var dlg = new ManualTargetInputForm())
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    // 입력값 유효성 검사
                    string name = dlg.TargetName?.Trim();
                    double ra = dlg.RA;
                    double dec = dlg.Dec;
                    double obsTime = dlg.ObsTime;

                    if (string.IsNullOrEmpty(name))
                    {
                        MessageBox.Show("대상 이름을 입력하세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    if (ra < 0 || ra >= 24)
                    {
                        MessageBox.Show("적경(RA)은 0~24 시간 범위로 입력하세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    if (dec < -90 || dec > 90)
                    {
                        MessageBox.Show("적위(Dec)는 -90~+90도 범위로 입력하세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    if (obsTime <= 0)
                    {
                        MessageBox.Show("관측 시간(분)은 0보다 커야 합니다.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // 중복 방지
                    if (observationTargets.Any(t => t.Name == name && Math.Abs(t.RA - ra) < 1e-6 && Math.Abs(t.Dec - dec) < 1e-6))
                    {
                        MessageBox.Show("이미 목록에 추가된 대상입니다.", "중복", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    observationTargets.Add(new ObservationTarget
                    {
                        Name = name,
                        RA = ra,
                        Dec = dec,
                        ObsTime = obsTime
                    });

                    // listBox2에 표시 (DataSource 재설정)
                    listBox2.DataSource = null;
                    listBox2.DataSource = observationTargets;
                    listBox2.DisplayMember = "ToString";
                }
            }
        }

        //============================================================================//


        // 불러오기 버튼
        private void button31_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "지평선 장애물 파일 (*.txt)|*.txt|모든 파일 (*.*)|*.*";
                openFileDialog.Title = "지평선 장애물 파일 불러오기";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedFilePath = openFileDialog.FileName;
                    // 파일 경로를 활용하여 장애물 점 리스트를 불러오고, obstacleAzList/obstacleAltList에 반영
                    try
                    {
                        var azList = new List<double>();
                        var altList = new List<double>();
                        using (var reader = new StreamReader(selectedFilePath, Encoding.UTF8))
                        {
                            string line;
                            while ((line = reader.ReadLine()) != null)
                            {
                                if (string.IsNullOrWhiteSpace(line)) continue;
                                if (line.StartsWith("#") || line.ToLower().Contains("az")) continue;
                                var parts = line.Split(new[] { ' ', '\t', ',' }, StringSplitOptions.RemoveEmptyEntries);
                                if (parts.Length < 2) continue;
                                if (double.TryParse(parts[0], out double az) && double.TryParse(parts[1], out double alt))
                                {
                                    azList.Add(az);
                                    altList.Add(alt);
                                }
                            }
                        }
                        obstacleAzList.Clear();
                        obstacleAltList.Clear();
                        obstacleAzList.AddRange(azList);
                        obstacleAltList.AddRange(altList);
                        isObstacleEditMode = true;
                        GenerateSkyPlot();
                        MessageBox.Show("지평선 장애물 파일을 불러왔습니다.", "불러오기 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"지평선 장애물 파일 읽기 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }


        // 장애물 제한 설정 버튼 실시간 점찍기)
        private void button4_Click(object sender, EventArgs e)
        {
            button13.Enabled = true;
            button14.Enabled = true;
            button15.Enabled = true;
            button16.Enabled = true;
            button18.Enabled = true;
            button20.Enabled = true;
            button21.Enabled = true;
            button27.Enabled = true;
            button28.Enabled = true;
            button29.Enabled = true;
            button30.Enabled = true;
            button32.Enabled = true;
            textBox8.Enabled = true;

            // Sky180Viewer 타이머 중지
            if (GenerateSkyPlotTimer != null)
                GenerateSkyPlotTimer.Stop();


            if (GenerateHorizonPlotTimer == null)
            {
                GenerateHorizonPlotTimer = new Timer();
                GenerateHorizonPlotTimer.Interval = 1000; // 1초
                GenerateHorizonPlotTimer.Tick += (s, ev) => DrawObstaclePlot();
            }
            GenerateHorizonPlotTimer.Start();

            // tabPage6가 열리도록 설정
            if (tabControl2 != null && tabPage6 != null)
            {
                tabControl2.SelectedTab = tabPage6;
            }

            // button31_Click(지평선 장애물 파일 불러오기)로 이미 obstacleAzList/obstacleAltList가 채워졌다면 그대로 사용
            // 아니라면 빈 리스트로 시작
            if (obstacleAzList == null || obstacleAltList == null || obstacleAzList.Count == 0 || obstacleAltList.Count == 0)
            {
                obstacleAzList.Clear();
                obstacleAltList.Clear();
                // horizon.txt은 신경쓰지 않음
            }
            isObstacleEditMode = true;
            DrawObstaclePlot();
        }


        // 점 찍기 버튼
        private void button21_Click(object sender, EventArgs e)
        {
            if (!isObstacleEditMode)
            {
                MessageBox.Show("장애물 편집 모드가 아닙니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (telescope == null || !telescope.Connected)
            {
                MessageBox.Show("가대가 연결되어 있지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                double az = telescope.Azimuth;
                double alt = telescope.Altitude;
                obstacleAzList.Add(az);
                obstacleAltList.Add(alt);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"가대 정보 읽기 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        // 되돌리기 버튼
        private void button20_Click(object sender, EventArgs e)
        {
            if (!isObstacleEditMode)
            {
                MessageBox.Show("장애물 편집 모드가 아닙니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (obstacleAzList.Count > 0)
            {
                obstacleAzList.RemoveAt(obstacleAzList.Count - 1);
                obstacleAltList.RemoveAt(obstacleAltList.Count - 1);
            }

        }


        // 장애물 점 그래프 그리기 (극좌표/2D)
        private void DrawObstaclePlot()
        {
            // 극좌표
            var plotModel = new OxyPlot.PlotModel
            {
                PlotAreaBorderColor = OxyPlot.OxyColors.Transparent,
                Background = OxyPlot.OxyColor.FromRgb(45, 55, 65),
                TextColor = OxyPlot.OxyColors.White,
            };
            var angleAxis = new OxyPlot.Axes.AngleAxis
            {
                Minimum = 0,
                Maximum = 360,
                MajorStep = 45,
                MinorStep = 15,
                StartAngle = 90,
                EndAngle = 450,
                FormatAsFractions = false,
                MajorGridlineStyle = OxyPlot.LineStyle.Solid,
                MinorGridlineStyle = OxyPlot.LineStyle.Dot,
                LabelFormatter = v =>
                {
                    switch ((int)v)
                    {
                        case 0: return "N";
                        case 90: return "E";
                        case 180: return "S";
                        case 270: return "W";
                        default: return v.ToString("0") + "°";
                    }
                },
                TicklineColor = OxyPlot.OxyColors.White,
                AxislineColor = OxyPlot.OxyColors.White,
                TextColor = OxyPlot.OxyColors.White,
                MajorGridlineColor = OxyPlot.OxyColors.White,
                MinorGridlineColor = OxyPlot.OxyColors.White
            };
            plotModel.Axes.Add(angleAxis);
            var radiusAxis = new OxyPlot.Axes.MagnitudeAxis
            {
                Minimum = 0,
                Maximum = 90,
                MajorStep = 30,
                MinorStep = 10,
                StartPosition = 0,
                EndPosition = 1,
                LabelFormatter = v => (90 - v).ToString("0") + "°",
                TicklineColor = OxyPlot.OxyColors.White,
                AxislineColor = OxyPlot.OxyColors.White,
                TextColor = OxyPlot.OxyColors.White,
                MajorGridlineColor = OxyPlot.OxyColors.White,
                MinorGridlineColor = OxyPlot.OxyColors.White
            };
            plotModel.Axes.Add(radiusAxis);

            // 장애물 곡선
            var series = new OxyPlot.Series.LineSeries
            {
                Color = OxyPlot.OxyColors.Orange,
                StrokeThickness = 2,
                Title = "지평선",
                MarkerType = OxyPlot.MarkerType.Circle,
                MarkerSize = 4,
                MarkerFill = OxyPlot.OxyColors.OrangeRed
            };
            for (int i = 0; i < obstacleAzList.Count; i++)
            {
                var point = new OxyPlot.DataPoint(90 - obstacleAltList[i], obstacleAzList[i]);
                series.Points.Add(point);
                plotModel.Annotations.Add(new OxyPlot.Annotations.TextAnnotation
                {
                    Text = $"{obstacleAzList[i]:0}°, {obstacleAltList[i]:0.#}°",
                    TextPosition = point,
                    Stroke = OxyPlot.OxyColors.Transparent,
                    TextColor = OxyPlot.OxyColors.Orange,
                    FontSize = 9,
                    Background = OxyPlot.OxyColor.FromAColor(120, OxyPlot.OxyColors.Black),
                    TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Left,
                    TextVerticalAlignment = OxyPlot.VerticalAlignment.Top
                });
            }
            plotModel.Series.Add(series);
            
            // ---- 가대 연결 시 현재 위치 노란색 점 표시 ----
            double? mountAz = null, mountAlt = null;
            try
            {
                if (telescope != null && telescope.Connected)
                {
                    mountAz = telescope.Azimuth;
                    mountAlt = telescope.Altitude;
                }
            }
            catch { /* 무시 */ }

            if (mountAz.HasValue && mountAlt.HasValue)
            {
                // 극좌표: (고도, 방위각)
                plotModel.Annotations.Add(new OxyPlot.Annotations.PointAnnotation
                {
                    X = 90 - mountAlt.Value, // OxyPlot 극좌표: 중심이 90, 가장자리가 0
                    Y = mountAz.Value,
                    Shape = OxyPlot.MarkerType.Circle,
                    Size = 10,
                    Fill = OxyPlot.OxyColors.Yellow,
                    Stroke = OxyPlot.OxyColors.Black,
                    StrokeThickness = 2,
                    ToolTip = "가대 현재 위치"
                });
            }
            
            // 내부/외부 영역 폴리곤
            if (obstacleAzList.Count > 1)
            {
                var areaPoints = new List<OxyPlot.DataPoint>();
                areaPoints.Add(new OxyPlot.DataPoint(90, 0));
                for (int i = 0; i < obstacleAzList.Count; i++)
                    areaPoints.Add(new OxyPlot.DataPoint(90 - obstacleAltList[i], obstacleAzList[i]));
                areaPoints.Add(new OxyPlot.DataPoint(90, 0));

                var polygonInner = new OxyPlot.Annotations.PolygonAnnotation
                {
                    Fill = OxyPlot.OxyColor.FromAColor(20, OxyPlot.OxyColors.LightYellow),
                    Stroke = OxyPlot.OxyColors.Transparent,
                };
                polygonInner.Points.AddRange(areaPoints);
                plotModel.Annotations.Add(polygonInner);

                var outerPoints = new List<OxyPlot.DataPoint>();
                for (int az = 0; az <= 360; az += 5)
                    outerPoints.Add(new OxyPlot.DataPoint(90, az));
                for (int i = areaPoints.Count - 1; i >= 0; i--)
                    outerPoints.Add(areaPoints[i]);
                var polygonOuter = new OxyPlot.Annotations.PolygonAnnotation
                {
                    Fill = OxyPlot.OxyColor.FromAColor(120, OxyPlot.OxyColors.DarkSlateBlue),
                    Stroke = OxyPlot.OxyColors.Transparent,
                };
                polygonOuter.Points.AddRange(outerPoints);
                plotModel.Annotations.Add(polygonOuter);
            }
            plotModel.Background = OxyPlot.OxyColor.FromRgb(20, 20, 30);

            // pictureBox4에 렌더링
            var pngExporter = new OxyPlot.WindowsForms.PngExporter
            {
                Width = pictureBox4.Width,
                Height = pictureBox4.Height
            };
            using (var stream = new MemoryStream())
            {
                pngExporter.Export(plotModel, stream);
                stream.Position = 0;
                pictureBox4.Image?.Dispose();
                pictureBox4.Image = Image.FromStream(stream);
            }

            // 2D 그래프
            if (pictureBox5 != null)
            {
                var plotModel2D = new OxyPlot.PlotModel
                {
                    PlotAreaBorderColor = OxyPlot.OxyColors.Transparent,
                    Background = OxyPlot.OxyColor.FromRgb(30, 30, 40),
                    TextColor = OxyPlot.OxyColors.White,
                };
                plotModel2D.Axes.Add(new OxyPlot.Axes.LinearAxis
                {
                    Position = OxyPlot.Axes.AxisPosition.Bottom,
                    Minimum = 0,
                    Maximum = 360,
                    MajorStep = 45,
                    MinorStep = 15,
                    Title = "방위각 (°)",
                    TextColor = OxyPlot.OxyColors.White,
                    AxislineColor = OxyPlot.OxyColors.White,
                    MajorGridlineStyle = OxyPlot.LineStyle.Solid,
                    MajorGridlineColor = OxyPlot.OxyColor.FromRgb(80, 90, 100),
                    FontSize = 11
                });
                plotModel2D.Axes.Add(new OxyPlot.Axes.LinearAxis
                {
                    Position = OxyPlot.Axes.AxisPosition.Left,
                    Minimum = 0,
                    Maximum = 90,
                    MajorStep = 30,
                    MinorStep = 10,
                    Title = "고도 (°)",
                    TextColor = OxyPlot.OxyColors.White,
                    AxislineColor = OxyPlot.OxyColors.White,
                    MajorGridlineStyle = OxyPlot.LineStyle.Solid,
                    MajorGridlineColor = OxyPlot.OxyColor.FromRgb(80, 90, 100),
                    FontSize = 11
                });
                var lineSeries2D = new OxyPlot.Series.LineSeries
                {
                    Color = OxyPlot.OxyColors.Orange,
                    StrokeThickness = 2,
                    Title = "지평선",
                    MarkerType = OxyPlot.MarkerType.Circle,
                    MarkerSize = 4,
                    MarkerFill = OxyPlot.OxyColors.OrangeRed
                };
                for (int i = 0; i < obstacleAzList.Count; i++)
                {
                    lineSeries2D.Points.Add(new OxyPlot.DataPoint(obstacleAzList[i], obstacleAltList[i]));
                }
                plotModel2D.Series.Add(lineSeries2D);
                var areaSeries = new OxyPlot.Series.AreaSeries
                {
                    Color = OxyPlot.OxyColor.FromAColor(120, OxyPlot.OxyColors.DarkSlateBlue),
                    MarkerType = OxyPlot.MarkerType.None
                };
                for (int i = 0; i < obstacleAzList.Count; i++)
                {
                    areaSeries.Points.Add(new OxyPlot.DataPoint(obstacleAzList[i], obstacleAltList[i]));
                    areaSeries.Points2.Add(new OxyPlot.DataPoint(obstacleAzList[i], 0));
                }
                // ---- 가대 연결 시 현재 위치 노란색 점 표시 (2D) ----
                if (mountAz.HasValue && mountAlt.HasValue)
                {
                    plotModel2D.Annotations.Add(new OxyPlot.Annotations.PointAnnotation
                    {
                        X = mountAz.Value,
                        Y = mountAlt.Value,
                        Shape = OxyPlot.MarkerType.Circle,
                        Size = 10,
                        Fill = OxyPlot.OxyColors.Yellow,
                        Stroke = OxyPlot.OxyColors.Black,
                        StrokeThickness = 2,
                        ToolTip = "가대 현재 위치"
                    });
                }
                plotModel2D.Series.Add(areaSeries);
                var pngExporter2D = new OxyPlot.WindowsForms.PngExporter
                {
                    Width = pictureBox5.Width,
                    Height = pictureBox5.Height
                };
                using (var stream2D = new MemoryStream())
                {
                    pngExporter2D.Export(plotModel2D, stream2D);
                    stream2D.Position = 0;
                    pictureBox5.Image?.Dispose();
                    pictureBox5.Image = Image.FromStream(stream2D);
                }
            }
        }


        // 장애물 제한 설정 해제 버튼
        private void button23_Click(object sender, EventArgs e)
        {
            button13.Enabled = false;
            button14.Enabled = false;
            button15.Enabled = false;
            button16.Enabled = false;
            button18.Enabled = false;
            button20.Enabled = false;
            button21.Enabled = false;
            button27.Enabled = false;
            button28.Enabled = false;
            button29.Enabled = false;
            button30.Enabled = false;
            button32.Enabled = false;
            textBox8.Enabled = false;

            if (GenerateSkyPlotTimer != null && GenerateSkyPlotTimer.Enabled == false)
            {
                GenerateSkyPlotTimer.Start();
            }

            if (GenerateHorizonPlotTimer != null)
                GenerateHorizonPlotTimer.Stop();

            GenerateSkyPlot();
        }


        // 장애물 제한 설정 저장 버튼
        private void button32_Click(object sender, EventArgs e)
        {
            if (obstacleAzList == null || obstacleAltList == null || obstacleAzList.Count == 0 || obstacleAzList.Count != obstacleAltList.Count)
            {
                MessageBox.Show("저장할 장애물 데이터가 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "지평선 장애물 파일 (*.txt)|*.txt|모든 파일 (*.*)|*.*";
                saveFileDialog.Title = "지평선 장애물 파일 저장";
                saveFileDialog.FileName = "horizon_obstacle.txt";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (var writer = new StreamWriter(saveFileDialog.FileName, false, Encoding.UTF8))
                        {
                            writer.WriteLine("# Az(deg)    Alt(deg)");
                            for (int i = 0; i < obstacleAzList.Count; i++)
                            {
                                writer.WriteLine($"{obstacleAzList[i]:F2}\t{obstacleAltList[i]:F2}");
                            }
                        }
                        MessageBox.Show("지평선 장애물 데이터가 저장되었습니다.", "저장 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"저장 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }


        // 방위각 0도, 고도 0도로 가대 이동 버튼
        private void button18_Click(object sender, EventArgs e)
        {
            try
            {
                if (telescope != null && telescope.Connected)
                {
                    // AlignmentMode을 통해 가대 타입 확인
                    // 0: AltAz, 1: Polar(적도의), 2: German Equatorial(적도의)
                    var alignmentMode = telescope.AlignmentMode;

                    // ... (생략) ...
                    if (alignmentMode == ASCOM.DeviceInterface.AlignmentModes.algPolar || alignmentMode == ASCOM.DeviceInterface.AlignmentModes.algGermanPolar) // 적도의(포크/독일식)
                    {
                        // 적도의: 방위각 0, 고도 0 → 적경/적위 변환
                        double lat = telescope.SiteLatitude;
                        double lon = telescope.SiteLongitude;
                        double alt = telescope.SiteElevation;
                        DateTime utcNow = DateTime.UtcNow;

                        // 방위각 0, 고도 0
                        double az = 0.0;
                        double altDeg = 0.0;

                        // 방위각/고도 → 적위
                        double latRad = lat * Math.PI / 180.0;
                        double altRad = altDeg * Math.PI / 180.0;
                        double azRad = az * Math.PI / 180.0;
                        double decRad = Math.Asin(Math.Sin(latRad) * Math.Sin(altRad) + Math.Cos(latRad) * Math.Cos(altRad) * Math.Cos(azRad));
                        double dec = decRad * 180.0 / Math.PI;

                        // 방위각/고도 → 시간각
                        double haRad = Math.Atan2(
                            -Math.Sin(azRad) * Math.Cos(altRad),
                            Math.Cos(latRad) * Math.Sin(altRad) - Math.Sin(latRad) * Math.Cos(altRad) * Math.Cos(azRad)
                        );
                        double ha = haRad * 180.0 / Math.PI;
                        if (ha < 0) ha += 360.0;
                        ha /= 15.0; // 시간각(시)

                        // LST 계산 (간이 구현, Astronomy.AstroTime 대체)
                        double jd = DateTimeToJulianDate(utcNow);
                        double lst = LocalSiderealTime(jd, lon); // (시)

                        // 적경 계산
                        double ra = lst - ha;
                        if (ra < 0) ra += 24.0;
                        if (ra > 24.0) ra -= 24.0;

                        // 이동 (적도의는 SlewToCoordinates 사용)
                        telescope.SlewToCoordinates(ra, dec);
                    }
                    else // AltAz
                    {
                        telescope.SlewToAltAz(0.0, 0.0);
                    }
                }
                else
                {
                    MessageBox.Show("가대가 연결되어 있지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"이동 명령 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // Julian Date 계산 (UTC 기준)
        private double DateTimeToJulianDate(DateTime date)
        {
            int Y = date.Year;
            int M = date.Month;
            double D = date.Day + (date.Hour + (date.Minute + date.Second / 60.0) / 60.0) / 24.0;

            if (M <= 2)
            {
                Y -= 1;
                M += 12;
            }
            int A = Y / 100;
            int B = 2 - A + (A / 4);
            double JD = Math.Floor(365.25 * (Y + 4716)) + Math.Floor(30.6001 * (M + 1)) + D + B - 1524.5;
            return JD;
        }


        // Local Sidereal Time 계산 (시 단위, 경도는 도 단위, 동경+)
        private double LocalSiderealTime(double jd, double longitude)
        {
            double T = (jd - 2451545.0) / 36525.0;
            double GMST = 280.46061837 + 360.98564736629 * (jd - 2451545.0)
                + 0.000387933 * T * T - T * T * T / 38710000.0;
            GMST = GMST % 360.0;
            if (GMST < 0) GMST += 360.0;
            double LMST = (GMST + longitude) % 360.0;
            if (LMST < 0) LMST += 360.0;
            return LMST / 15.0; // 시 단위
        }


        // 현재 가대 위치에서 방위각 N도, 고도 M도 만큼 이동한 위치의 (적경, 적위)를 반환하는 함수
        // N: 방위각 변화량(도), M: 고도 변화량(도)
        // 반환값: (RA, Dec) (RA: 시간 단위, Dec: 도 단위)
        public (double RA, double Dec) GetEquatorialAfterAltAzOffset(double deltaAz, double deltaAlt)
        {
            if (telescope == null || !telescope.Connected)
                throw new InvalidOperationException("가대가 연결되어 있지 않습니다.");

            // 1. 현재 방위각/고도 읽기
            double currentAz = telescope.Azimuth;
            double currentAlt = telescope.Altitude;

            // 2. 이동 후 방위각/고도 계산 (0~360, 0~90 범위로 정규화)
            double newAz = currentAz + deltaAz;
            if (newAz < 0) newAz += 360;
            if (newAz >= 360) newAz -= 360;
            double newAlt = currentAlt + deltaAlt;
            if (newAlt > 90) newAlt = 90;
            if (newAlt < 0) newAlt = 0;

            // 3. 현재 시간, 관측지 위도/경도
            DateTime now = DateTime.Now;
            var site = GetObservingSite();
            double lat = site.lat;
            double lon = site.lon;

            // 4. 방위각/고도 → 적경/적위 변환
            // (1) 고도/방위각 → 적위
            double latRad = lat * Math.PI / 180.0;
            double altRad = newAlt * Math.PI / 180.0;
            double azRad = newAz * Math.PI / 180.0;
            double decRad = Math.Asin(Math.Sin(latRad) * Math.Sin(altRad) + Math.Cos(latRad) * Math.Cos(altRad) * Math.Cos(azRad));
            double dec = decRad * 180.0 / Math.PI;

            // (2) 고도/방위각 → 시간각
            double haRad = Math.Atan2(
                -Math.Sin(azRad) * Math.Cos(altRad),
                Math.Cos(latRad) * Math.Sin(altRad) - Math.Sin(latRad) * Math.Cos(altRad) * Math.Cos(azRad)
            );
            double ha = haRad * 180.0 / Math.PI;
            if (ha < 0) ha += 360.0;
            ha /= 15.0; // 시간각(시)

            // (3) LST 계산
            double jd = DateTimeToJulianDate(now.ToUniversalTime());
            double lst = LocalSiderealTime(jd, lon); // (시)

            // (4) 적경 계산
            double ra = lst - ha;
            if (ra < 0) ra += 24.0;
            if (ra > 24.0) ra -= 24.0;

            return (ra, dec);
        }


        // 고도 +N° 이동 버튼 (AltAz/Equatorial에 따라 다르게 동작)
        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                if (telescope != null && telescope.Connected)
                {
                    double moveDeg = 15.0;
                    if (double.TryParse(textBox8.Text, out double userDeg) && userDeg > 0)
                        moveDeg = Math.Abs(userDeg);

                    var alignmentMode = telescope.AlignmentMode;
                    if (alignmentMode == ASCOM.DeviceInterface.AlignmentModes.algAltAz)
                    {
                        // AltAz: 고도 +N도
                        double currentAz = telescope.Azimuth;
                        double currentAlt = telescope.Altitude;
                        double newAlt = currentAlt + moveDeg;
                        // 접근 불가 범위 제한 (예: 고도 0~90도만 허용)
                        if (newAlt > 90) newAlt = 90;
                        if (newAlt < 0) newAlt = 0;
                        // 추가: 극단적 각도(예: 0~90도 이외)로 이동 시도시 경고
                        if (newAlt == 0 || newAlt == 90)
                        {
                            MessageBox.Show("가대가 접근할 수 없는 극단적 고도(0° 또는 90°)로 이동이 제한됩니다.", "이동 제한", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        telescope.SlewToAltAz(currentAz, newAlt);
                    }
                    else // 적도의(포크/독일식/독일식)
                    {
                        // 현재 위치에서 moveDeg 만큼 고도를 이동한 후의 적경/적위를 계산
                        var (ra, dec) = GetEquatorialAfterAltAzOffset(0, moveDeg);

                        telescope.SlewToCoordinates(ra, dec);
                    }
                }
                else
                {
                    MessageBox.Show("가대가 연결되어 있지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"이동 명령 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // 고도 -N° 이동 버튼 (AltAz/Equatorial에 따라 다르게 동작)
        private void button14_Click(object sender, EventArgs e)
        {
            try
            {
                if (telescope != null && telescope.Connected)
                {
                    double moveDeg = 15.0;
                    if (double.TryParse(textBox8.Text, out double userDeg) && userDeg > 0)
                        moveDeg = Math.Abs(userDeg);

                    var alignmentMode = telescope.AlignmentMode;
                    if (alignmentMode == ASCOM.DeviceInterface.AlignmentModes.algAltAz)
                    {
                        // AltAz: 고도 -N도
                        double currentAz = telescope.Azimuth;
                        double currentAlt = telescope.Altitude;
                        double newAlt = currentAlt - moveDeg;
                        // 접근 불가 범위 제한
                        if (newAlt > 90) newAlt = 90;
                        if (newAlt < 0) newAlt = 0;
                        if (newAlt == 0 || newAlt == 90)
                        {
                            MessageBox.Show("가대가 접근할 수 없는 극단적 고도(0° 또는 90°)로 이동이 제한됩니다.", "이동 제한", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        telescope.SlewToAltAz(currentAz, newAlt);
                    }
                    else // 적도의(포크/독일식/독일식)
                    {
                        // 현재 위치에서 moveDeg 만큼 고도를 이동한 후의 적경/적위를 계산
                        var (ra, dec) = GetEquatorialAfterAltAzOffset(0, -moveDeg);

                        telescope.SlewToCoordinates(ra, dec);
                    }
                }
                else
                {
                    MessageBox.Show("가대가 연결되어 있지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"이동 명령 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // 방위각 +N° 이동 버튼 (AltAz/Equatorial에 따라 다르게 동작)
        private void button15_Click(object sender, EventArgs e)
        {
            try
            {
                if (telescope != null && telescope.Connected)
                {
                    double moveDeg = 15.0;
                    if (double.TryParse(textBox8.Text, out double userDeg) && userDeg > 0)
                        moveDeg = Math.Abs(userDeg);

                    var alignmentMode = telescope.AlignmentMode;
                    if (alignmentMode == ASCOM.DeviceInterface.AlignmentModes.algAltAz)
                    {
                        // AltAz: 방위각 +N도
                        double currentAz = telescope.Azimuth;
                        double currentAlt = telescope.Altitude;
                        double newAz = currentAz + moveDeg;
                        // 접근 불가 범위 제한 (0~360도)
                        if (newAz >= 360) newAz -= 360;
                        if (newAz < 0) newAz += 360;
                        // 고도 극단값(0, 90)에서 이동 시도시 경고
                        if (currentAlt == 0 || currentAlt == 90)
                        {
                            MessageBox.Show("가대가 접근할 수 없는 극단적 고도(0° 또는 90°)에서의 이동이 제한됩니다.", "이동 제한", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        telescope.SlewToAltAz(newAz, currentAlt);
                    }
                    else // 적도의(포크/독일식/독일식)
                    {
                        // 현재 위치에서 moveDeg 만큼 방위각을 이동한 후의 적경/적위를 계산
                        var (ra, dec) = GetEquatorialAfterAltAzOffset(moveDeg, 0);

                        telescope.SlewToCoordinates(ra, dec);
                    }
                }
                else
                {
                    MessageBox.Show("가대가 연결되어 있지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"이동 명령 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // 방위각 -N° 이동 버튼 (AltAz/Equatorial에 따라 다르게 동작)
        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                if (telescope != null && telescope.Connected)
                {
                    double moveDeg = 15.0;
                    if (double.TryParse(textBox8.Text, out double userDeg) && userDeg > 0)
                        moveDeg = Math.Abs(userDeg);

                    var alignmentMode = telescope.AlignmentMode;
                    if (alignmentMode == ASCOM.DeviceInterface.AlignmentModes.algAltAz)
                    {
                        // AltAz: 방위각 -N도
                        double currentAz = telescope.Azimuth;
                        double currentAlt = telescope.Altitude;
                        double newAz = currentAz - moveDeg;
                        if (newAz >= 360) newAz -= 360;
                        if (newAz < 0) newAz += 360;
                        if (currentAlt == 0 || currentAlt == 90)
                        {
                            MessageBox.Show("가대가 접근할 수 없는 극단적 고도(0° 또는 90°)에서의 이동이 제한됩니다.", "이동 제한", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        telescope.SlewToAltAz(newAz, currentAlt);           
                    }
                    else // 적도의(포크/독일식/독일식)
                    {
                        // 현재 위치에서 -moveDeg 만큼 방위각을 이동한 후의 적경/적위를 계산
                        var (ra, dec) = GetEquatorialAfterAltAzOffset(-moveDeg, 0);

                        telescope.SlewToCoordinates(ra, dec);
                    }
                }
                else
                {
                    MessageBox.Show("가대가 연결되어 있지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"이동 명령 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // 방위각 +N°,  고도 +N°이동 버튼 (AltAz/Equatorial에 따라 다르게 동작)
        private void button27_Click(object sender, EventArgs e)
        {
            try
            {
                if (telescope != null && telescope.Connected)
                {
                    double moveDeg = 15.0;
                    if (double.TryParse(textBox8.Text, out double userDeg) && userDeg > 0)
                        moveDeg = Math.Abs(userDeg);

                    var alignmentMode = telescope.AlignmentMode;
                    if (alignmentMode == ASCOM.DeviceInterface.AlignmentModes.algAltAz)
                    {
                        // AltAz: 방위각 +N도, 고도 +N도
                        double currentAz = telescope.Azimuth;
                        double currentAlt = telescope.Altitude;
                        double newAz = currentAz + moveDeg;
                        double newAlt = currentAlt + moveDeg;
                        if (newAz >= 360) newAz -= 360;
                        if (newAz < 0) newAz += 360;
                        if (newAlt > 90) newAlt = 90;
                        if (newAlt < 0) newAlt = 0;
                        if (newAlt == 0 || newAlt == 90)
                        {
                            MessageBox.Show("가대가 접근할 수 없는 극단적 고도(0° 또는 90°)로 이동이 제한됩니다.", "이동 제한", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        telescope.SlewToAltAz(newAz, newAlt);
                    }
                    else
                    {
                        var (ra, dec) = GetEquatorialAfterAltAzOffset(moveDeg, moveDeg);

                        telescope.SlewToCoordinates(ra, dec);
                    }
                }
                else
                {
                    MessageBox.Show("가대가 연결되어 있지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"이동 명령 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // 방위각 -N°,  고도 +N°이동 버튼 (AltAz/Equatorial에 따라 다르게 동작)
        private void button30_Click(object sender, EventArgs e)
        {
            try
            {
                if (telescope != null && telescope.Connected)
                {
                    double moveDeg = 15.0;
                    if (double.TryParse(textBox8.Text, out double userDeg) && userDeg > 0)
                        moveDeg = Math.Abs(userDeg);

                    var alignmentMode = telescope.AlignmentMode;
                    if (alignmentMode == ASCOM.DeviceInterface.AlignmentModes.algAltAz)
                    {
                        // AltAz: 방위각 -N도, 고도 +N도
                        double currentAz = telescope.Azimuth;
                        double currentAlt = telescope.Altitude;
                        double newAz = currentAz - moveDeg;
                        double newAlt = currentAlt + moveDeg;
                        if (newAz >= 360) newAz -= 360;
                        if (newAz < 0) newAz += 360;
                        if (newAlt > 90) newAlt = 90;
                        if (newAlt < 0) newAlt = 0;
                        if (newAlt == 0 || newAlt == 90)
                        {
                            MessageBox.Show("가대가 접근할 수 없는 극단적 고도(0° 또는 90°)로 이동이 제한됩니다.", "이동 제한", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        telescope.SlewToAltAz(newAz, newAlt);
                    }
                    else
                    {
                        var (ra, dec) = GetEquatorialAfterAltAzOffset(-moveDeg, moveDeg);

                        telescope.SlewToCoordinates(ra, dec);
                    }
                }
                else
                {
                    MessageBox.Show("가대가 연결되어 있지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"이동 명령 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // 방위각 -N°,  고도 -N°이동 버튼 (AltAz/Equatorial에 따라 다르게 동작)
        private void button29_Click(object sender, EventArgs e)
        {
            try
            {
                if (telescope != null && telescope.Connected)
                {
                    double moveDeg = 15.0;
                    if (double.TryParse(textBox8.Text, out double userDeg) && userDeg > 0)
                        moveDeg = Math.Abs(userDeg);

                    var alignmentMode = telescope.AlignmentMode;
                    if (alignmentMode == ASCOM.DeviceInterface.AlignmentModes.algAltAz)
                    {
                        // AltAz: 방위각 -N도, 고도 -N도
                        double currentAz = telescope.Azimuth;
                        double currentAlt = telescope.Altitude;
                        double newAz = currentAz - moveDeg;
                        double newAlt = currentAlt - moveDeg;
                        if (newAz >= 360) newAz -= 360;
                        if (newAz < 0) newAz += 360;
                        if (newAlt > 90) newAlt = 90;
                        if (newAlt < 0) newAlt = 0;
                        if (newAlt == 0 || newAlt == 90)
                        {
                            MessageBox.Show("가대가 접근할 수 없는 극단적 고도(0° 또는 90°)로 이동이 제한됩니다.", "이동 제한", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        telescope.SlewToAltAz(newAz, newAlt);
                    }
                    else
                    {
                        var (ra, dec) = GetEquatorialAfterAltAzOffset(-moveDeg, -moveDeg);

                        telescope.SlewToCoordinates(ra, dec);
                    }
                }
                else
                {
                    MessageBox.Show("가대가 연결되어 있지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"이동 명령 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // 방위각 +N°,  고도 -N°이동 버튼 (AltAz/Equatorial에 따라 다르게 동작)
        private void button28_Click(object sender, EventArgs e)
        {
            try
            {
                if (telescope != null && telescope.Connected)
                {
                    double moveDeg = 15.0;
                    if (double.TryParse(textBox8.Text, out double userDeg) && userDeg > 0)
                        moveDeg = Math.Abs(userDeg);

                    var alignmentMode = telescope.AlignmentMode;
                    if (alignmentMode == ASCOM.DeviceInterface.AlignmentModes.algAltAz)
                    {
                        // AltAz: 방위각 +N도, 고도 -N도
                        double currentAz = telescope.Azimuth;
                        double currentAlt = telescope.Altitude;
                        double newAz = currentAz + moveDeg;
                        double newAlt = currentAlt - moveDeg;
                        if (newAz >= 360) newAz -= 360;
                        if (newAz < 0) newAz += 360;
                        if (newAlt > 90) newAlt = 90;
                        if (newAlt < 0) newAlt = 0;
                        if (newAlt == 0 || newAlt == 90)
                        {
                            MessageBox.Show("가대가 접근할 수 없는 극단적 고도(0° 또는 90°)로 이동이 제한됩니다.", "이동 제한", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        telescope.SlewToAltAz(newAz, newAlt);
                    }
                    else
                    {
                        var (ra, dec) = GetEquatorialAfterAltAzOffset(moveDeg, -moveDeg);

                        telescope.SlewToCoordinates(ra, dec);
                    }
                }
                else
                {
                    MessageBox.Show("가대가 연결되어 있지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"이동 명령 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // 장애물 데이터 접근을 위한 public 메서드들 추가
        public List<double> GetObstacleAzList()
        {
            return obstacleAzList;
        }

        public List<double> GetObstacleAltList()
        {
            return obstacleAltList;
        }


        // 경로 계산하기 버튼
        // 여기 부분 확인 필요!!!!!!!!!!!!!!!!!!!!!!!(2025.08.15 박병민)
        private void button5_Click(object sender, EventArgs e)
        {
            // 경로 계산 전 타이머 중지
            if (GenerateSkyPlotTimer != null && GenerateSkyPlotTimer.Enabled)
                GenerateSkyPlotTimer.Stop();
            if (GenerateHorizonPlotTimer != null && GenerateHorizonPlotTimer.Enabled)
                GenerateHorizonPlotTimer.Stop();

            // 경로 계산 전 망원경 연결 해제
            bool wasTelescopeConnected = telescope != null && telescope.Connected;
            if (wasTelescopeConnected)
            {
                try
                {
                    telescope.Connected = false;
                }
                catch { }
            }

            CalculateObservationRoute();  // 테스트 실행

            // 경로 계산 후 망원경 연결 복구
            if (wasTelescopeConnected)
            {
                try
                {
                    telescope.Connected = true;
                }
                catch { }
            }

            GenerateSkyPlot();

            // 경로 계산 후 타이머 재시작
            if (GenerateSkyPlotTimer != null && (GenerateHorizonPlotTimer == null || !GenerateHorizonPlotTimer.Enabled))
                GenerateSkyPlotTimer.Start();
            if (GenerateHorizonPlotTimer != null && GenerateHorizonPlotTimer.Enabled)
                GenerateHorizonPlotTimer.Start();
        }


        // 경로 저장하기 버튼
        private void button7_Click(object sender, EventArgs e)
        {
            if (finalObservationOrder.Count == 0)
            {
                MessageBox.Show("저장할 관측 경로가 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "관측 경로 (*.txt)|*.txt|모든 파일 (*.*)|*.*";
                saveFileDialog.Title = "관측 경로 저장";
                saveFileDialog.FileName = "observation_path.txt";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (var writer = new StreamWriter(saveFileDialog.FileName, false, Encoding.UTF8))
                        {
                            // 헤더
                            writer.WriteLine("#Name|RA|Dec|ObsTime(min)");
                            foreach (var target in finalObservationOrder)
                            {
                                // 구분자: | (파이프)
                                writer.WriteLine($"{target.Name}|{target.RA}|{target.Dec}|{target.ObsTime}");
                            }
                        }
                        MessageBox.Show("관측 경로가 저장되었습니다.", "저장 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"저장 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }


        // 관측 경로 불러오기 버튼
        private void button8_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "관측 경로 (*.txt)|*.txt|모든 파일 (*.*)|*.*";
                openFileDialog.Title = "관측 경로  불러오기";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var loadedTargets = new List<ObservationTarget>();
                        using (var reader = new StreamReader(openFileDialog.FileName, Encoding.UTF8))
                        {
                            string line;
                            while ((line = reader.ReadLine()) != null)
                            {
                                if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#")) continue;
                                var parts = line.Split('|');
                                if (parts.Length < 4) continue;

                                string name = parts[0];
                                if (!double.TryParse(parts[1], out double ra)) continue;
                                if (!double.TryParse(parts[2], out double dec)) continue;
                                if (!double.TryParse(parts[3], out double obsTime)) obsTime = 10;

                                loadedTargets.Add(new ObservationTarget
                                {
                                    Name = name,
                                    RA = ra,
                                    Dec = dec,
                                    ObsTime = obsTime
                                });
                            }
                        }
                        // finalObservationOrder만 갱신 (관측 경로 불러오기이므로)
                        // 결과 출력 (listBox3에 출력)
                        StringBuilder sb = new StringBuilder();

                        finalObservationOrder.Clear();
                        finalObservationOrder.AddRange(loadedTargets);
                        // listBox3에만 반영 (관측 경로 전용)

                        foreach (var t in finalObservationOrder)
                        {
                            sb.AppendLine($"{t.Name}, {t.ObsTime}분 관측");
                        }

                        listBox3.DisplayMember = "ToString";

                        // listBox3에 출력
                        if (listBox3 != null)
                        {
                            listBox3.Items.Clear();
                            foreach (var line in sb.ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.None))
                            {
                                if (!string.IsNullOrWhiteSpace(line))
                                    listBox3.Items.Add(line);
                            }
                        }
                        MessageBox.Show("관측 경로를 불러왔습니다.", "불러오기 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"불러오기 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }


        // 관측 시작 버튼
        private async void button6_Click(object sender, EventArgs e)
        {
            if (telescope == null || !telescope.Connected)
            {
                MessageBox.Show("가대가 연결되어 있지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (finalObservationOrder == null || finalObservationOrder.Count == 0)
            {
                MessageBox.Show("관측 대상 목록이 비어 있습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 순차적으로 이동 및 관측
            for (int i = 0; i < finalObservationOrder.Count; i++)
            {
                var target = finalObservationOrder[i];
                double ra = target.RA;
                double dec = target.Dec;
                double obsTimeMin = target.ObsTime;

                // 가대 이동중 표시
                if (label17 != null)
                {
                    label17.Text = $"[{target.Name}]로 가대 이동중...";
                    label17.ForeColor = Color.Orange;
                    label17.Visible = true;
                    // UI 갱신
                    label17.Refresh();
                    Application.DoEvents();
                }

                // 이동
                try
                {
                    telescope.SlewToCoordinates(ra, dec);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"[{target.Name}]로 이동 중 오류: {ex.Message}", "이동 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                }

                // 관측중 표시
                if (label17 != null)
                {
                    label17.Text = $"[{target.Name}] 관측중...";
                    label17.ForeColor = Color.Green;
                    label17.Visible = true;
                    label17.Refresh();
                    Application.DoEvents();
                }

                int obsTimeMs = (int)(obsTimeMin * 60 * 1000);
                int interval = 1000;
                int waited = 0;
                while (waited < obsTimeMs)
                {
                    await System.Threading.Tasks.Task.Delay(interval);
                    Application.DoEvents();
                    waited += interval;
                }
            }

            telescope.Tracking = false; // 관측 완료 후 추적 해제

            if (label17 != null)
            {
                label17.Text = "모든 관측 대상에 대한 관측이 완료되었습니다.";
                label17.ForeColor = Color.Blue;
                label17.Visible = true;
                label17.Refresh();
            }

            MessageBox.Show("모든 관측 대상에 대한 관측이 완료되었습니다.", "관측 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        // 폼 닫을 때 정리
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DisconnectTelescope();
        }


        //===========================================================================//
        // ========================= 미구현 (필요없음) ==============================//


        private void label7_Click(object sender, EventArgs e)
        {

        }
        private void label13_Click(object sender, EventArgs e)
        {

        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }
        private void tabPage4_Click(object sender, EventArgs e)
        {

        }
        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }
        private void label18_Click(object sender, EventArgs e)
        {

        }

        //============================================================================//
        //============================================================================//


        // OptimizedRouteFinder 클래스에 있는 경로 최적화 알고리즘을 실행하는 메서드
        // 클래스 필드로 최종 관측 순서 및 정보 리스트 추가
        private List<ObservationTarget> finalObservationOrder = new List<ObservationTarget>();

        public void CalculateObservationRoute()
        {
            // 관측 대상이 없으면 리턴
            if (observationTargets == null || observationTargets.Count == 0)
            {
                MessageBox.Show("관측 대상 목록이 비어 있습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 실전 데이터 준비
            List<int> idxs = Enumerable.Range(0, observationTargets.Count).ToList();
            List<double> RAs = observationTargets.Select(t => t.RA).ToList();
            List<double> Decs = observationTargets.Select(t => t.Dec).ToList();
            List<double> ObsTimes = observationTargets.Select(t => t.ObsTime).ToList();

            DateTime start = GetSelectedDateTime();
            var sw = new System.Diagnostics.Stopwatch();
            sw.Start();
            
            // 정보 전달 잘 되는 것 확인 (2025.08.11 박병민)
            List<int> route = finder.FindRoute(idxs, RAs, Decs, ObsTimes, start);

            // 경로가 비어있는 경우!!!!! ->> 관측 가능 대상 없음!!!!!
            if (route == null || route.Count == 0)
            {
                MessageBox.Show("관측 가능한 천체가 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            sw.Stop();
            double elapsed = sw.Elapsed.TotalMilliseconds;

            // 최종 관측 순서 및 정보 리스트 생성
            finalObservationOrder = route.Select(i => new ObservationTarget
            {
                Name = observationTargets[i].Name,
                RA = observationTargets[i].RA,
                Dec = observationTargets[i].Dec,
                ObsTime = observationTargets[i].ObsTime
            }).ToList();

            // 결과 출력 (listBox3에 출력)
            StringBuilder sb = new StringBuilder();
            foreach (var t in finalObservationOrder)
            {
                sb.AppendLine($"{t.Name}, {t.ObsTime}분 관측");
            }
            label14.Text = $"경로 계산 시간: {elapsed:F2} ms";

            label16.Text = "경로 탐색 완료!";
            label16.ForeColor = Color.Green;
            label16.Visible = false;

            // 0.1초 대기 (UI 응답 유지, WinForms에서는 Application.DoEvents와 Timer/Task.Delay 사용)
            // label16을 두 번 깜빡이게 구현
            var timer = new Timer();
            int blinkCount = 0;
            timer.Interval = 100; // 0.1초 = 100ms
            timer.Tick += (s, ev) =>
            {
                label16.Visible = !label16.Visible;
                blinkCount++;
                if (blinkCount >= 4) // 두 번 깜빡임 (on-off-on-off)
                {
                    timer.Stop();
                    label16.Visible = true; // 마지막엔 보이게
                }
            };
            timer.Start();
            Application.DoEvents();

            // listBox3에 출력
            if (listBox3 != null)
            {
                listBox3.Items.Clear();
                foreach (var line in sb.ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.None))
                {
                    if (!string.IsNullOrWhiteSpace(line))
                        listBox3.Items.Add(line);
                }
            }
        }
        

        // textBox7 시간 자동 갱신 및 수동 입력 모드 관리
        private Timer textBox7Timer;


        // textBox7의 시간 자동 갱신 타이머 시작
        private void StartTextBox7Timer()
        {
            if (textBox7Timer == null)
            {
                textBox7Timer = new Timer();
                textBox7Timer.Interval = 1000;
                textBox7Timer.Tick += (s, ev) =>
                {
                    // 날짜 부분은 그대로 두고, 시간만 현재로 갱신
                    string currentText = textBox7.Text;
                    string datePart = DateTime.Now.ToString("yyyy-MM-dd");
                    string timePart = DateTime.Now.ToString("HH:mm:ss");

                    if (!string.IsNullOrWhiteSpace(currentText) && currentText.Length >= 10)
                    {
                        string oldDatePart = currentText.Substring(0, 10);
                        if (oldDatePart[4] == '-' && oldDatePart[7] == '-')
                            datePart = oldDatePart;
                    }

                    textBox7.Text = $"{datePart} {timePart}";
                };
            }
            if (!textBox7Timer.Enabled)
                textBox7Timer.Start();
        }


        // textBox7의 시간 자동 갱신 타이머 중지
        private void StopTextBox7Timer()
        {
            if (textBox7Timer != null && textBox7Timer.Enabled)
                textBox7Timer.Stop();
        }


        // textBox7_TextChanged: radioButton3이 체크되면 타이머 시작, 아니면 중지
        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            if (radioButton3 != null && radioButton3.Checked)
                StartTextBox7Timer();
            else
                StopTextBox7Timer();
        }


        // radioButton3: "현재 시간 자동" 모드
        private void radioButton3_CheckedChanged_1(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                textBox7.ReadOnly = true;
                StartTextBox7Timer();
            }
            else
            {
                StopTextBox7Timer();
            }
        }


        // radioButton4: "수동 입력" 모드
        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked)
            {
                StopTextBox7Timer();
                textBox7.ReadOnly = true;
                textBox7.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                textBox7.MaxLength = 19;
                textBox7.Click -= textBox7_Click; // 중복 방지
                textBox7.Click += textBox7_Click;
            }
            else
            {
                textBox7.Click -= textBox7_Click;
                textBox7.ReadOnly = true;
            }
        }


        // textBox7 클릭 시 시간 입력 폼을 띄움 (수동 입력 모드에서만)
        private void textBox7_Click(object sender, EventArgs e)
        {
            if (!radioButton4.Checked) return;
            DateTime initial = DateTime.Now;
            DateTime.TryParse(textBox7.Text, out initial);
            using (var dlg = new TimeInputForm(initial))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    textBox7.Text = dlg.SelectedDateTime.ToString("yyyy-MM-dd HH:mm:ss");
                }
            }
        }


        // 텍스트박스7에서 시간 가져오는 유틸리티 (radioButton3/4 모드에 따라)
        private DateTime GetSelectedDateTime()
        {
            if (radioButton3 != null && radioButton3.Checked)
            {
                return DateTime.Now;
            }
            else if (radioButton4 != null && radioButton4.Checked)
            {
                DateTime dt;
                if (DateTime.TryParse(textBox7.Text, out dt))
                    return dt;
                else
                    return DateTime.Now;
            }
            else
            {
                return DateTime.Now;
            }
        }


        // 맨 위로 버튼
        private void button12_Click(object sender, EventArgs e)
        {
            int selectedIndex = listBox3.SelectedIndex;
            if (selectedIndex > 0 && selectedIndex < finalObservationOrder.Count)
            {
                // 관측 경로 리스트에서 선택된 항목을 맨 위로 이동
                var item = finalObservationOrder[selectedIndex];
                finalObservationOrder.RemoveAt(selectedIndex);
                finalObservationOrder.Insert(0, item);

                // listBox3 갱신
                listBox3.Items.Clear();
                foreach (var t in finalObservationOrder)
                {
                    listBox3.Items.Add($"{t.Name}, {t.ObsTime}분 관측");
                }
                listBox3.SelectedIndex = 0;
            }
        }


        // 위로 버튼
        private void button17_Click(object sender, EventArgs e)
        {
            int selectedIndex = listBox3.SelectedIndex;
            if (selectedIndex > 0 && selectedIndex < finalObservationOrder.Count)
            {
                // 관측 경로 리스트에서 선택된 항목을 한 칸 위로 이동
                var item = finalObservationOrder[selectedIndex];
                finalObservationOrder.RemoveAt(selectedIndex);
                finalObservationOrder.Insert(selectedIndex - 1, item);

                // listBox3 갱신
                listBox3.Items.Clear();
                foreach (var t in finalObservationOrder)
                {
                    listBox3.Items.Add($"{t.Name}, {t.ObsTime}분 관측");
                }
                listBox3.SelectedIndex = selectedIndex - 1;
            }
        }


        // 아래로 버튼
        private void button19_Click(object sender, EventArgs e)
        {
            int selectedIndex = listBox3.SelectedIndex;
            if (selectedIndex >= 0 && selectedIndex < finalObservationOrder.Count - 1)
            {
                // 관측 경로 리스트에서 선택된 항목을 한 칸 아래로 이동
                var item = finalObservationOrder[selectedIndex];
                finalObservationOrder.RemoveAt(selectedIndex);
                finalObservationOrder.Insert(selectedIndex + 1, item);

                // listBox3 갱신
                listBox3.Items.Clear();
                foreach (var t in finalObservationOrder)
                {
                    listBox3.Items.Add($"{t.Name}, {t.ObsTime}분 관측");
                }
                listBox3.SelectedIndex = selectedIndex + 1;
            }
        }


        // 맨 아래로 버튼
        private void button22_Click(object sender, EventArgs e)
        {
            int selectedIndex = listBox3.SelectedIndex;
            if (selectedIndex >= 0 && selectedIndex < finalObservationOrder.Count - 1)
            {
                // 관측 경로 리스트에서 선택된 항목을 맨 아래로 이동
                var item = finalObservationOrder[selectedIndex];
                finalObservationOrder.RemoveAt(selectedIndex);
                finalObservationOrder.Add(item);

                // listBox3 갱신
                listBox3.Items.Clear();
                foreach (var t in finalObservationOrder)
                {
                    listBox3.Items.Add($"{t.Name}, {t.ObsTime}분 관측");
                }
                listBox3.SelectedIndex = finalObservationOrder.Count - 1;
            }
        }


        // 가대를 움직여서 가대의 적경축/적위축의 최대 속도와 가속도를 자동으로 구하는 버튼
        private async void button33_Click(object sender, EventArgs e)
        {
            if (telescope == null || !telescope.Connected)
            {
                MessageBox.Show("가대가 연결되어 있지 않습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 측정 결과를 저장할 변수
            double maxSpeedRA = 0, maxSpeedDec = 0;
            double maxAccelRA = 0, maxAccelDec = 0;

            try
            {
                // 1. 현재 위치 저장
                double startRA = telescope.RightAscension;
                double startDec = telescope.Declination;

                // 2. RA축 최대 속도 측정
                double testDeltaRA = 3.0; // 1시간각(=15도) 이동
                double targetRA = startRA + testDeltaRA;
                if (targetRA > 24) targetRA -= 24;

                DateTime t0 = DateTime.Now;
                telescope.SlewToCoordinates(targetRA, startDec);
                while (telescope.Slewing)
                {
                    await System.Threading.Tasks.Task.Delay(50);
                }
                DateTime t1 = DateTime.Now;
                double dtRA = (t1 - t0).TotalSeconds;
                double movedRADeg = Math.Abs((targetRA - startRA) * 15.0);
                if (movedRADeg > 180) movedRADeg = 360 - movedRADeg;
                maxSpeedRA = movedRADeg / dtRA;

                // 3. Dec축 최대 속도 측정
                double testDeltaDec = 30.0; // 10도 이동
                double targetDec = startDec + testDeltaDec;
                if (targetDec > 90) targetDec = 90;
                if (targetDec < -90) targetDec = -90;

                t0 = DateTime.Now;
                telescope.SlewToCoordinates(targetRA, targetDec);
                while (telescope.Slewing)
                {
                    await System.Threading.Tasks.Task.Delay(50);
                }
                t1 = DateTime.Now;
                double dtDec = (t1 - t0).TotalSeconds;
                double movedDecDeg = Math.Abs(targetDec - startDec);
                maxSpeedDec = movedDecDeg / dtDec;

                // 4. RA축 가속도 근사 측정
                double shortDeltaRA = 1; // 1시간각(=15도)
                double shortTargetRA = startRA + shortDeltaRA;
                if (shortTargetRA > 24) shortTargetRA -= 24;

                t0 = DateTime.Now;
                telescope.SlewToCoordinates(shortTargetRA, targetDec);
                while (telescope.Slewing)
                {
                    await System.Threading.Tasks.Task.Delay(50);
                }
                t1 = DateTime.Now;
                double dtShortRA = (t1 - t0).TotalSeconds;
                double movedShortRADeg = Math.Abs((shortTargetRA - targetRA) * 15.0);
                if (movedShortRADeg > 180) movedShortRADeg = 360 - movedShortRADeg;
                // 등가속도 공식: s = 0.5*a*t^2 -> a = 2s/t^2
                maxAccelRA = (dtShortRA > 0) ? 2 * movedShortRADeg / (dtShortRA * dtShortRA) : 0;

                // 5. Dec축 가속도 근사 측정
                double shortDeltaDec = 15.0; // 1도
                double shortTargetDec = targetDec + shortDeltaDec;
                if (shortTargetDec > 90) shortTargetDec = 90;
                if (shortTargetDec < -90) shortTargetDec = -90;

                t0 = DateTime.Now;
                telescope.SlewToCoordinates(shortTargetRA, shortTargetDec);
                while (telescope.Slewing)
                {
                    await System.Threading.Tasks.Task.Delay(50);
                }
                t1 = DateTime.Now;
                double dtShortDec = (t1 - t0).TotalSeconds;
                double movedShortDecDeg = Math.Abs(shortTargetDec - targetDec);
                maxAccelDec = (dtShortDec > 0) ? 2 * movedShortDecDeg / (dtShortDec * dtShortDec) : 0;

                // 6. 결과 표시
                MessageBox.Show(
                    $"RA축 최대 속도: {maxSpeedRA:F2} 도/초\n" +
                    $"Dec축 최대 속도: {maxSpeedDec:F2} 도/초\n" +
                    $"RA축 최대 가속도: {maxAccelRA:F2} 도/초²\n" +
                    $"Dec축 최대 가속도: {maxAccelDec:F2} 도/초²",
                    "가대 속도/가속도 측정 결과", MessageBoxButtons.OK, MessageBoxIcon.Information
                );

                // 측정값을 각 텍스트박스에 표시
                textBox9.Text = maxSpeedRA.ToString("F2");
                textBox10.Text = maxSpeedDec.ToString("F2");
                textBox11.Text = maxAccelRA.ToString("F2");
                textBox12.Text = maxAccelDec.ToString("F2");

                // 7. 원래 위치로 복귀
                telescope.SlewToCoordinates(startRA, startDec);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"측정 중 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Form1 클래스의 textBox9, textBox10, textBox11, textBox12 필드가 private으로 선언되어 있기 때문에
        // 외부 클래스에서 직접 접근할 수 없음
        // 이는 C#의 접근 제한자(캡슐화) 원칙에 따름
        // 해결방법 -> Form1에 public/protected internal getter 메서드(속성)를 추가하여 값을 간접적으로 읽을 수 있게 한다

        public double MotorSpeedRAValue
        {
            get
            {
                double v;
                if (double.TryParse(textBox9.Text, out v) && v > 0) return v;
                return 0;
            }
        }
        public double MotorSpeedDecValue
        {
            get
            {
                double v;
                if (double.TryParse(textBox10.Text, out v) && v > 0) return v;
                return 0;
            }
        }
        public double MotorAccelerationRAValue
        {
            get
            {
                double v;
                if (double.TryParse(textBox11.Text, out v) && v > 0) return v;
                return 0;
            }
        }
        public double MotorAccelerationDecValue
        {
            get
            {
                double v;
                if (double.TryParse(textBox12.Text, out v) && v > 0) return v;
                return 0;
            }
        }
    }


    public class OptimizedRouteFinder
    {
        // Form1에서 정의된 함수를 사용하기 위한 과정..
        private Form1 form;

        // ASCOM Telescope 객체 사용
        private ASCOM.DriverAccess.Telescope telescope;

        // ASCOM Telescope 객체를 선택적으로 받아 초기화
        public OptimizedRouteFinder(Form1 form, ASCOM.DriverAccess.Telescope telescope = null)
        {
            this.form = form;
            this.telescope = telescope;
        }

        // 가대 설정 메서드
        public void SetTelescope(ASCOM.DriverAccess.Telescope telescope)
        {
            this.telescope = telescope;
        }


        // 1) 두 점 간 각거리 계산 (deg)
        // RA/Dec 리스트가 '시간(hour), 도' 형태라면 AngularDistance 진입 전에 RA를 도 단위로 변환
        double AngularDistance(double ra1Hour, double dec1Deg, double ra2Hour, double dec2Deg)
        {
            double ra1 = ra1Hour * 15.0;
            double ra2 = ra2Hour * 15.0;

            double radRa1 = ra1 * Math.PI / 180.0;
            double radDec1 = dec1Deg * Math.PI / 180.0;
            double radRa2 = ra2 * Math.PI / 180.0;
            double radDec2 = dec2Deg * Math.PI / 180.0;

            double cosD = Math.Sin(radDec1) * Math.Sin(radDec2) +
                          Math.Cos(radDec1) * Math.Cos(radDec2) * Math.Cos(radRa1 - radRa2);

            return Math.Acos(Math.Min(1.0, Math.Max(-1.0, cosD))) * 180.0 / Math.PI;
        }


        // 2) 각 점마다 k번째 최근접 거리 계산 (k-distance)
        List<double> ComputeKDistance(List<double> RAs, List<double> Decs, int k)
        {
            int N = RAs.Count;
            var kDistances = new List<double>(N);

            for (int i = 0; i < N; i++)
            {
                List<double> dists = new List<double>();
                for (int j = 0; j < N; j++)
                {
                    if (i == j) continue;
                    double dist = AngularDistance(RAs[i], Decs[i], RAs[j], Decs[j]);
                    dists.Add(dist);
                }
                dists.Sort();
                // 기존 코드
                // kDistances.Add(dists[k - 1]); // k번째 가까운 거리 (k=1이면 1번째 최근접)

                // 수정된 코드 (k-1이 dists.Count보다 크면 마지막 값 사용, dists가 비어있으면 0)
                if (dists.Count == 0)
                {
                    kDistances.Add(0);
                }
                else if (k - 1 < dists.Count)
                {
                    kDistances.Add(dists[k - 1]);
                }
                else
                {
                    kDistances.Add(dists[dists.Count - 1]);
                }
            }
            return kDistances;
        }


        // 3) 꺾임점(Elbow) 자동 탐지 - 기존: 단순 2차 미분 최소
        double FindElbowThreshold(List<double> sortedDistances)
        {
            // (이전 로직은 남겨두되, 새 DetermineClusterRadius 에서 사용하지 않을 수도 있음)
            int n = sortedDistances.Count;
            if (n < 3) return (n > 0 ? sortedDistances[n - 1] : 0);

            double minSecondDiff = double.MaxValue;
            int elbowIndex = 1;
            for (int i = 1; i < n - 1; i++)
            {
                double secondDiff = Math.Abs(sortedDistances[i - 1] - 2 * sortedDistances[i] + sortedDistances[i + 1]);
                if (secondDiff < minSecondDiff)
                {
                    minSecondDiff = secondDiff;
                    elbowIndex = i;
                }
            }
            return sortedDistances[elbowIndex];
        }

        // --- 추가: k-distance 기반 적응형 클러스터 반경 결정 (r) ---
        double DetermineClusterRadius(List<double> kDistances, int totalCount,
                                      int desiredAvgClusterSize = 5,
                                      double hardMinDeg = 0.2,
                                      double hardMaxDeg = 15.0)
        {
            // 1. 유효 데이터 정리
            var clean = kDistances.Where(d => d > 0).OrderBy(d => d).ToList();
            if (clean.Count == 0) return 1.5; // fallback

            double median = clean[clean.Count / 2];
            double p40 = clean[(int)(clean.Count * 0.40)];
            double p60 = clean[(int)(clean.Count * 0.60)];
            double p80 = clean[(int)(clean.Count * 0.80)];

            // 2. Kneedle 방식 (끝점 연결 직선 대비 최대 양의 편차)
            // 점 (i, clean[i]) 와 (0, clean[0]), (n-1, clean[n-1]) 이용
            double first = clean.First();
            double last = clean.Last();
            int n = clean.Count;
            double maxDev = double.MinValue;
            double kneedleValue = median; // 초기값

            for (int i = 0; i < n; i++)
            {
                // 직선 보간 값
                double t = (double)i / (n - 1);
                double lineVal = first + (last - first) * t;
                double dev = clean[i] - lineVal;
                if (dev > maxDev)
                {
                    maxDev = dev;
                    kneedleValue = clean[i];
                }
            }

            // 3. 기본 후보 r
            // - Kneedle 값이 median 보다 너무 작으면 p60 사용
            double rCandidate = kneedleValue;
            if (rCandidate < median * 0.9)
                rCandidate = p60;

            // 4. 범위 제한 (median 기반 스케일)
            double lowerBound = Math.Max(hardMinDeg, median * 0.8);
            double upperBound = Math.Min(hardMaxDeg, median * 2.2);
            rCandidate = Math.Min(Math.Max(rCandidate, lowerBound), upperBound);

            // 5. 첫 번째 품질 점검: 예상 평균 클러스터 크기 추정이 어려우므로 사후 조정은 FindSubRoute 내부에서 실제 클러스터 수를 본 뒤 반복
            return rCandidate;
        }


        // 4) 밀도기반 클러스터링 (DBSCAN 유사)
        List<List<int>> ClusterByAngularDistance(List<int> idxs, List<double> RAs, List<double> Decs, double maxDistanceDeg)
        {
            int N = idxs.Count;
            int[] clusterIds = new int[N];
            for (int i = 0; i < N; i++) clusterIds[i] = -1;

            int clusterId = 0;
            for (int i = 0; i < N; i++)
            {
                if (clusterIds[i] != -1) continue;
                clusterIds[i] = clusterId;
                var queue = new Queue<int>();
                queue.Enqueue(i);

                while (queue.Count > 0)
                {
                    int current = queue.Dequeue();
                    for (int j = 0; j < N; j++)
                    {
                        if (clusterIds[j] != -1) continue;
                        double dist = AngularDistance(RAs[current], Decs[current], RAs[j], Decs[j]);
                        if (dist <= maxDistanceDeg)
                        {
                            clusterIds[j] = clusterId;
                            queue.Enqueue(j);
                        }
                    }
                }
                clusterId++;
            }

            var clusters = new List<List<int>>();
            for (int c = 0; c < clusterId; c++)
                clusters.Add(new List<int>());

            for (int i = 0; i < N; i++)
                clusters[clusterIds[i]].Add(idxs[i]);

            return clusters;
        }


        // 5) 순열 생성기 (yield return)
        IEnumerable<List<T>> Permute<T>(List<T> list, int start, int end)
        {
            if (start == end)
                yield return new List<T>(list);
            else
            {
                for (int i = start; i < end; i++)
                {
                    (list[start], list[i]) = (list[i], list[start]);
                    foreach (var perm in Permute(list, start + 1, end))
                        yield return perm;
                    (list[start], list[i]) = (list[i], list[start]);
                }
            }
        }


        // 6) 클러스터 내부 최적 경로 brute-force 탐색
        (List<int>, double) SolveWithinCluster(List<int> clusterIdxs, List<double> RAs, List<double> Decs, DateTime starttime,
                                               Func<double, double, DateTime, double> HA_cal,
                                               Func<double, double, double, double, DateTime, double> MovingTime)
        {
            int N = clusterIdxs.Count;
            if (N == 0) return (new List<int>(), 0.0);
            if (N == 1) return (new List<int> { clusterIdxs[0] }, 0.0);

            double[,] distances = new double[N, N];
            for (int i = 0; i < N; i++)
                for (int j = 0; j < N; j++)
                    distances[i, j] = MovingTime(RAs[i], Decs[i], RAs[j], Decs[j], starttime);

            double minLen = double.MaxValue;
            List<int> bestPath = null;

            foreach (var perm in Permute(clusterIdxs, 0, N))
            {
                double len = 0;
                for (int i = 0; i < N - 1; i++)
                {
                    int fromIdx = clusterIdxs.IndexOf(perm[i]);
                    int toIdx = clusterIdxs.IndexOf(perm[i + 1]);
                    len += distances[fromIdx, toIdx];
                }
                if (len < minLen)
                {
                    minLen = len;
                    bestPath = new List<int>(perm);
                }
            }
            return (bestPath ?? new List<int>(), minLen);
        }

        // 7) 최종 통합 FindSubRoute 함수 --> 2025.09.05 기준 최적경로 찾지 못함문제 발견!!!

        // 7) 최종 통합 FindSubRoute 함수 (r 적응 개선)
        (List<int>, double) FindSubRoute(
            List<int> idxs,
            List<double> RAs,
            List<double> Decs,
            DateTime starttime,
            Func<double, double, DateTime, double> HA_cal,
            Func<double, double, double, double, DateTime, double> MovingTime,
            int k_for_kdist = 2)
        {
            int N = idxs.Count;
            if (N == 0) return (new List<int>(), 0.0);
            if (N == 1) return (new List<int> { idxs[0] }, 0.0);

            // --- k-거리 계산 및 정렬
            var kDistances = ComputeKDistance(RAs, Decs, k_for_kdist);
            kDistances.Sort();

            // --- 1차 반경 산출
            double maxClusterAngularDistDeg = DetermineClusterRadius(kDistances, N);

            // --- 반복 조정: 클러스터가 1개이거나 평균 크기 과대/과소일 때 반경 shrink/expand
            List<List<int>> clusters = null;
            int iteration = 0;
            while (iteration < 8)
            {
                clusters = ClusterByAngularDistance(idxs, RAs, Decs, maxClusterAngularDistDeg);
                int cCount = clusters.Count;
                if (cCount == 0) break;

                double avgSize = clusters.Average(c => c.Count);

                bool tooManyMerged = (cCount == 1 && N > 4) || avgSize > 2 * 5; // 목표 평균≈5 가정
                bool tooFragmented = avgSize < 2 && cCount > N / 2;             // 거의 모두 단일/쪼개짐

                if (!tooManyMerged && !tooFragmented) break;

                if (tooManyMerged)
                {
                    // 반경 축소
                    maxClusterAngularDistDeg *= 0.75;
                    if (maxClusterAngularDistDeg < 0.2) { maxClusterAngularDistDeg = 0.2; break; }
                }
                else if (tooFragmented)
                {
                    // 반경 확대 (단, 과도 확대 방지)
                    maxClusterAngularDistDeg *= 1.25;
                    if (maxClusterAngularDistDeg > 15.0) { maxClusterAngularDistDeg = 15.0; break; }
                }
                iteration++;
            }

            // 최종 클러스터 없으면 fallback: 모든 대상 brute-force
            if (clusters == null || clusters.Count == 0)
            {
                var fullClusterRAs = idxs.Select(id => RAs[idxs.IndexOf(id)]).ToList();
                var fullClusterDecs = idxs.Select(id => Decs[idxs.IndexOf(id)]).ToList();
                return SolveWithinCluster(idxs, fullClusterRAs, fullClusterDecs, starttime, HA_cal, MovingTime);
            }

            // --- 클러스터 대표 좌표 (중심)
            var clusterCenters = clusters.Select(cluster =>
            {
                var ras = cluster.Select(id => RAs[idxs.IndexOf(id)]).ToList();
                var decs = cluster.Select(id => Decs[idxs.IndexOf(id)]).ToList();
                return (Idxs: cluster, RA: ras.Average(), Dec: decs.Average());
            }).ToList();

            int cN = clusterCenters.Count;
            if (cN == 0) return (new List<int>(), 0.0);
            if (cN == 1)
            {
                var clusterRAs = clusterCenters[0].Idxs.Select(id => RAs[idxs.IndexOf(id)]).ToList();
                var clusterDecs = clusterCenters[0].Idxs.Select(id => Decs[idxs.IndexOf(id)]).ToList();
                return SolveWithinCluster(clusterCenters[0].Idxs, clusterRAs, clusterDecs, starttime, HA_cal, MovingTime);
            }

            // --- cluster 대표 HA 계산 (마지막 cluster = HA 최소 → 기존 로직 유지하되 flip 감소 위해 HA 정렬 재검토)
            double[] HAs = new double[cN];
            double min_HA = double.MaxValue;
            int min_HA_idx = 0;
            for (int i = 0; i < cN; i++)
            {
                HAs[i] = HA_cal(clusterCenters[i].RA, clusterCenters[i].Dec, starttime);
                if (HAs[i] < min_HA)
                {
                    min_HA = HAs[i];
                    min_HA_idx = i;
                }
            }

            // --- cluster 간 이동시간
            double[,] clusterDistances = new double[cN, cN];
            for (int i = 0; i < cN; i++)
                for (int j = 0; j < cN; j++)
                    clusterDistances[i, j] = MovingTime(clusterCenters[i].RA, clusterCenters[i].Dec,
                                                        clusterCenters[j].RA, clusterCenters[j].Dec,
                                                        starttime);

            // --- 순열 탐색 (마지막 cluster 고정)
            var clusterOrderCandidates = new List<int>();
            for (int i = 0; i < cN; i++)
                if (i != min_HA_idx) clusterOrderCandidates.Add(i);

            double minPathLen = double.MaxValue;
            List<int> bestClusterOrder = null;

            foreach (var perm in Permute(clusterOrderCandidates, 0, clusterOrderCandidates.Count))
            {
                double len = 0;
                for (int i = 0; i < perm.Count - 1; i++)
                    len += clusterDistances[perm[i], perm[i + 1]];

                if (perm.Count > 0)
                    len += clusterDistances[perm[perm.Count - 1], min_HA_idx];

                if (len < minPathLen)
                {
                    minPathLen = len;
                    bestClusterOrder = new List<int>(perm);
                }
            }
            if (bestClusterOrder == null) bestClusterOrder = new List<int>();

            // --- 최종 경로 구성 
            var finalRoute = new List<int>();
            double totalLength = 0;
            DateTime currentTime = starttime;
            int? prevLastIdx = null;

            foreach (var cIdx in bestClusterOrder)
            {
                var cluster = clusterCenters[cIdx];
                var clusterRAs = cluster.Idxs.Select(id => RAs[idxs.IndexOf(id)]).ToList();
                var clusterDecs = cluster.Idxs.Select(id => Decs[idxs.IndexOf(id)]).ToList();
                var (intraPath, intraLen) = SolveWithinCluster(cluster.Idxs, clusterRAs, clusterDecs, currentTime, HA_cal, MovingTime);

                if (prevLastIdx != null && intraPath.Count > 0)
                {
                    int fromIdx = prevLastIdx.Value;
                    int toIdx = intraPath.First();
                    double moveLen = MovingTime(RAs[idxs.IndexOf(fromIdx)], Decs[idxs.IndexOf(fromIdx)],
                                                RAs[idxs.IndexOf(toIdx)], Decs[idxs.IndexOf(toIdx)], currentTime);
                    totalLength += moveLen;
                    currentTime = currentTime.AddSeconds(moveLen);
                }

                finalRoute.AddRange(intraPath);
                totalLength += intraLen;
                currentTime = currentTime.AddSeconds(intraLen);
                if (intraPath.Count > 0) prevLastIdx = intraPath.Last();
            }

            // 마지막 cluster
            {
                var lastCluster = clusterCenters[min_HA_idx];
                var clusterRAs = lastCluster.Idxs.Select(id => RAs[idxs.IndexOf(id)]).ToList();
                var clusterDecs = lastCluster.Idxs.Select(id => Decs[idxs.IndexOf(id)]).ToList();
                var (intraPath, intraLen) = SolveWithinCluster(lastCluster.Idxs, clusterRAs, clusterDecs, currentTime, HA_cal, MovingTime);

                if (prevLastIdx != null && intraPath.Count > 0)
                {
                    int fromIdx = prevLastIdx.Value;
                    int toIdx = intraPath.First();
                    double moveLen = MovingTime(RAs[idxs.IndexOf(fromIdx)], Decs[idxs.IndexOf(fromIdx)],
                                                RAs[idxs.IndexOf(toIdx)], Decs[idxs.IndexOf(toIdx)], currentTime);
                    totalLength += moveLen;
                    currentTime = currentTime.AddSeconds(moveLen);
                }

                finalRoute.AddRange(intraPath);
                totalLength += intraLen;
            }

            return (finalRoute, totalLength);
        }



        // 경로 탐색 메서드
        public List<int> FindRoute(List<int> idxs, List<double> RAs, List<double> Decs, List<double> ObsTimes, DateTime starttime)
        {
            DateTime currenttime = starttime;
            int N = idxs.Count;
            List<int> SolIdxs = new List<int>();

            int epoch = 0; //epoch for break
            while (true)
            {
                // Console.WriteLine($"Epoch: {epoch}, CurrentTime: {currenttime}");

                // SubRoute를 탐색할 subgroup 생성
                List<int> subgroup_idxs = new List<int>();
                List<double> subgroup_RAs = new List<double>();
                List<double> subgroup_Decs = new List<double>();
                List<bool> candidates_temp = TemporalCandidates(RAs, Decs, ObsTimes, currenttime, SolIdxs);

                for (int i = 0; i < N; i++)
                {
                    if (candidates_temp[i])
                    {
                        subgroup_idxs.Add(idxs[i]);
                        subgroup_RAs.Add(RAs[i]);
                        subgroup_Decs.Add(Decs[i]);
                    }
                }

                // 기존 코드 (DBSCAN 사용하기 전)
                //(List<int> subroute_temp, double subroutelength) = FindSubRoute(subgroup_idxs, subgroup_RAs, subgroup_Decs, currenttime);

                // 수정된 코드: FindSubRoute에 필요한 델리게이트 전달 -> DBSCAN 사용!! --> 2025.09.05 기준 최적경로 찾지 못함문제 발견!!!
                (List<int> subroute_temp, double subroutelength) = FindSubRoute(
                    subgroup_idxs,
                    subgroup_RAs,
                    subgroup_Decs,
                    currenttime,
                    (ra, dec, t) => HA_cal(ra, dec, t),
                    (ra1, dec1, ra2, dec2, t) => MovingTime(ra1, dec1, ra2, dec2, t)
                );
                SolIdxs.AddRange(subroute_temp);

                // Console.WriteLine($"Subroute length: {subroutelength}, Subroute count: {subroute_temp.Count}");

                currenttime = currenttime.AddSeconds(subroutelength); // 움직이는데 걸리는 시간
                for (int i = 0; i < subroute_temp.Count; i++)
                    currenttime = currenttime.AddMinutes(ObsTimes[subroute_temp[i]]); //총 노출시간
                // ----------------------------------------------------
                // 무한 루프 방지
                // ----------------------------------------------------
                // Console.WriteLine($"Updated CurrentTime: {currenttime}");

                if (SolIdxs.Count == N || epoch >= 100) break;
                epoch++;
            }

            return SolIdxs;
        }


        // 주어진 천체의 관측 가능 여부를 종합하여 판단하는 메서드들
        private List<bool> TemporalCandidates(List<double> RAs, List<double> Decs, List<double> ObsTimes, DateTime time, List<int> DoneIdxs)
        {
            // [주어신 시각에 관측 가능한 천체들 후보 리스트(bool) 반환]
            // RAs, Decs : 모든 천체들의 좌표
            // time : 현재 시간
            // DoneIdxs : 이미 관측해서 고려 안해도 될 놈들

            List<bool> Candidates = new List<bool>();
            for (int i = 0; i < RAs.Count; i++) Candidates.Add(true);

            for (int i = 0; i < RAs.Count; i++)
            {
                bool rise = RiseObservable(RAs[i], Decs[i], time);

                int minutes = (int)ObsTimes[i];
                int seconds = (int)((ObsTimes[i] - minutes) * 60);
                DateTime settime = time.Add(new TimeSpan(0, 0, minutes, seconds));

                bool set = RiseObservable(RAs[i], Decs[i], settime);
                bool cloud = CloudObservable(RAs[i], Decs[i], time);
                bool obstacle = ObstacleObservable(RAs[i], Decs[i], time);
                bool moonlight = MoonlightObservable(RAs[i], Decs[i], time);

                // 디버깅 코드 추가
                // System.Diagnostics.Debug.WriteLine($"[TemporalCandidates] idx: {i}, RA: {RAs[i]}, Dec: {Decs[i]}, Time: {time}, Rise: {rise}, Cloud: {cloud}, Obstacle: {obstacle}, Done: {DoneIdxs.Contains(i)}");

                Candidates[i] &= rise;
                Candidates[i] &= set;
                Candidates[i] &= cloud;
                Candidates[i] &= obstacle;

                if (DoneIdxs.Contains(i)) Candidates[i] = false;
            }

            return Candidates;
        }


        private (List<int>, double) FindSubRoute(List<int> idxs, List<double> RAs, List<double> Decs, DateTime starttime)
        {
            // [주어진 후보군 내 최적의 경로를 Brute-Force(노가다)로 찾기]
            // 시간각이 가장 뒤쪽에 있는 천체는 도착지점으로 고정.
            // idxs : subgroup(subroute를 탐색할 천체들의 인덱스)
            // RAs, Decs : subgroup의 좌표들
            // starttime : subgroup 탐색 시작 시간

            int N = idxs.Count;

            // 빈 subgroup 처리
            if (N == 0)
            {
                return (new List<int>(), 0.0);
            }

            // 천체가 1개뿐인 경우 처리
            if (N == 1)
            {
                return (new List<int> { idxs[0] }, 0.0);
            }

            double[] HAs = new double[N];
            double[,] distances = new double[N, N];

            // 시간각 리스트 계산 & 마지막 천체 (시간각이 가장 뒤쪽에 있는 천체) 탐색
            double min_HA = double.MaxValue;
            int min_HA_idx = 0;
            for (int i = 0; i < N; i++)
            {
                HAs[i] = HA_cal(RAs[i], Decs[i], starttime);
                if (HAs[i] < min_HA)
                {
                    min_HA = HAs[i];
                    min_HA_idx = i;
                }
            }

            // 거리-인접 행렬 계산
            for (int i = 0; i < N; i++)
            {
                for (int j = 0; j < N; j++)
                {
                    distances[i, j] = MovingTime(RAs[i], Decs[i], RAs[j], Decs[j], starttime);
                }
            }

            // Brute-Force 위한 순열을 할 대상들 (마지막 도착지 제외 전부)
            List<int> SubRouteObjects = new List<int>();
            for (int i = 0; i < idxs.Count; i++)
            {
                if (i == min_HA_idx) continue;
                SubRouteObjects.Add(idxs[i]);
            }

            // SubRouteObjects 내용 출력
            System.Diagnostics.Debug.WriteLine("SubRouteObjects: " + string.Join(", ", SubRouteObjects));

            // SubRouteObjects가 비어있는 경우 (천체가 1개뿐인 경우)
            if (SubRouteObjects.Count == 0)
            {
                List<int> singleRoute = new List<int> { idxs[min_HA_idx] };
                return (singleRoute, 0.0);
            }

            // 순열을 한 번에 모두 메모리에 저장하지 않고, yield로 하나씩 처리
            // 각 순열별(경로별) pathlength 계산 & 최단 경로 탐색
            double min_pathlength = double.MaxValue;
            List<int> min_path = null;

            foreach (var perm in Permute(SubRouteObjects, 0, SubRouteObjects.Count))
            {
                // 현재 순열 내용 출력
                System.Diagnostics.Debug.WriteLine("perm: " + string.Join(", ", perm));

                double pathlength_temp = 0;

                // 순열 내부 경로 길이 계산
                for (int j = 0; j < perm.Count - 1; j++)
                {
                    // idxs 배열에서의 인덱스를 찾아서 사용
                    int from_idx = Array.IndexOf(idxs.ToArray(), perm[j]);
                    int to_idx = Array.IndexOf(idxs.ToArray(), perm[j + 1]);
                    pathlength_temp += distances[from_idx, to_idx];
                }

                // 마지막 천체로의 이동 시간 추가
                if (perm.Count > 0)
                {
                    int last_idx = Array.IndexOf(idxs.ToArray(), perm[perm.Count - 1]);
                    pathlength_temp += distances[last_idx, min_HA_idx];
                }

                // 현재 순열의 pathlength 출력
                System.Diagnostics.Debug.WriteLine($"perm pathlength: {pathlength_temp}");

                if (pathlength_temp < min_pathlength)
                {
                    min_pathlength = pathlength_temp;
                    min_path = new List<int>(perm);
                }
            }

            if (min_path == null) min_path = new List<int>(); // subroute의 해답
            min_path.Add(idxs[min_HA_idx]); // 마지막 천체를 subroute에 추가해주기
            return (min_path, min_pathlength);
        }


        // IEnumerable 기반 순열 생성기 (메모리 절약)
        private IEnumerable<List<int>> Permute(List<int> list, int k, int n)
        {
            if (k == n)
            {
                yield return new List<int>(list);
            }
            else
            {
                for (int i = k; i < n; i++)
                {
                    Swap(list, k, i);
                    foreach (var perm in Permute(list, k + 1, n))
                        yield return perm;
                    Swap(list, k, i);
                }
            }
        }


        // 모든 순열을 메모리에 저장하는 기존 방식
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


        // 리스트의 두 요소를 스왑하는 유틸리티 메서드
        private void Swap<T>(List<T> list, int i, int j)
        {
            T temp = list[i];
            list[i] = list[j];
            list[j] = temp;
        }


        // 천체가 관측 가능한지 여부를 판단하는 메서드
        private bool RiseObservable(double RA, double Dec, DateTime time)
        {
            var site = form.GetObservingSite();
            var equator = new Equator
            {
                RA = RA * 15, // RA를 시간 단위에서 도 단위로 변환!!!!!!
                Dec = Dec
            };
            double altitude = CoordinateSystem.GetElevationAngle(time, equator, latitude: site.lat, longitude: site.lon);

            // 디버깅 코드 추가
            // System.Diagnostics.Debug.WriteLine($"[RiseObservable] RA: {RA}, Dec: {Dec}, Time: {time}, Lat: {site.lat}, Lon: {site.lon}, Altitude: {altitude}");

            return altitude > 0;
        }

        // 천체가 구름으로로 인해 관측 가능한지 여부를 판단하는 메서드
        private bool CloudObservable(double RA, double Dec, DateTime time) => true;


        // 천체가 장애물로 인해 관측 가능한지 여부를 판단하는 메서드
        private bool ObstacleObservable(double RA, double Dec, DateTime time)
        {
            // Form1에서 장애물 데이터 가져오기
            var obstacleAzList = form.GetObstacleAzList();
            var obstacleAltList = form.GetObstacleAltList();

            // 장애물 데이터가 없거나 비어있으면 관측 가능
            if (obstacleAzList == null || obstacleAltList == null ||
                obstacleAzList.Count < 2 || obstacleAzList.Count != obstacleAltList.Count)
            {
                return true;
            }

            // 천체의 현재 지평 좌표 계산
            var site = form.GetObservingSite();
            var equator = new Equator
            {
                RA = RA * 15, // RA를 시간 단위에서 도 단위로 변환
                Dec = Dec
            };

            double targetAltitude = CoordinateSystem.GetElevationAngle(time, equator, latitude: site.lat, longitude: site.lon);
            double targetAzimuth = CoordinateSystem.GetAzimuth(time, equator, latitude: site.lat, longitude: site.lon);

            // 고도가 0도 미만이면 지평선 아래이므로 관측 불가
            if (targetAltitude <= 0)
            {
                return false;
            }

            // 방위각을 0~360도 범위로 정규화
            while (targetAzimuth < 0) targetAzimuth += 360;
            while (targetAzimuth >= 360) targetAzimuth -= 360;

            // 해당 방위각에서의 장애물 고도를 선형 보간으로 계산
            double obstacleAltitudeAtAzimuth = InterpolateObstacleAltitude(targetAzimuth, obstacleAzList, obstacleAltList);

            // 천체의 고도가 장애물 고도보다 높으면 관측 가능
            return targetAltitude > obstacleAltitudeAtAzimuth;
        }


        // ----------- 25.08.31~25.09.03 수정 -----------
        // 천체가 달빛으로 인해 관측 가능한지 여부를 판단하는 메서드
        private bool MoonlightObservable(double RA, double Dec, DateTime time)
        {   
            (double moonRA, double moonDec) = GetMoonEquatorialCoordinates(time); // 달의 적경, 적위 계산

            double obsZenithDist = ZenithDistance(RA, Dec, time); // 천체의 천정거리 계산
            double moonZenithDist = ZenithDistance(moonRA, moonDec, time); // 달의 천정거리 계산
            double theta = AngularDistance(RA, Dec, moonRA, moonDec); //angularSeperation // 천체와 달 사이의 각거리 계산
            double FL0 = MoonFlux0(time);
            double moonlitIntensity = MoonlitIntensity(obsZenithDist, moonZenithDist, theta, FL0); // 달빛 산란 강도 계산

            //Debug.WriteLine($"FL0: {FL0}");
            Debug.WriteLine($"MoonMag: {-2.5 * (Math.Log10(FL0) - (Math.Log10(3631) - 26))}");

            return SkySurfaceBrightnessDetector(moonlitIntensity);// 현재는 하늘 표면밝기 기준 정하여 구현 (다른 기준으로 수정 가능)
        }

        private bool SkySurfaceBrightnessDetector(double moonlitIntensity)
        {
            bool result;
            double eps = 1e-25; // 0 대신 넣을 아주 작은 값
            double SkyThreshold = 20;// mag/arcsec2기준 표면밝기 (수정 가능)
            if (moonlitIntensity <= 0) moonlitIntensity = eps;

            double SkySurfaceBrightness = -2.5 * (Math.Log10(moonlitIntensity) - 2*Math.Log10(206265) - (Math.Log10(3631)-26));
            Debug.WriteLine($"SkySurfaceBrightness: {SkySurfaceBrightness}");
            result = SkySurfaceBrightness > SkyThreshold;
            return result;
        }
        
        private double MoonFlux0(DateTime time, double lambda=0.55)
        {
            // lambda : 관측 파장 (μm)
            double T_sun = 5600; //K, 태양 유효온도
            double R_sun = 696000; // km, 태양의 반지름

            double R_moon = 1737.4; // km, 달의 반지름
            double distance_sun_earth = 149600000; //km, 지구와 태양 사이의 평균 거리
            double distance_earth_moon = 384400; // km, 지구와 달 사이의 평균 거리

            double logSunFlux = logPlanckBB(lambda, T_sun);// 태양 방출 플럭스 (by Blackbody Radiation)
            double S_sun = Math.Pow(10, Math.Log10(Math.PI) + logSunFlux + 2 * Math.Log10(R_sun/distance_sun_earth));
            
            (double moonRA, double moonDec) = GetMoonEquatorialCoordinates(time); // 달의 적경, 적위 계산
            (double sunRA, double sunDec) = GetSunEquatorialCoordinates(time); // 태양의 적경, 적위 계산
            double alpha = AngularDistance(moonRA, moonDec, sunRA, sunDec); // 위상각 (radian) (0도 : 삭, 180도 : 망)

            double albedo_moon = AlbdedoMoon(alpha);

            double FL0 = S_sun * (R_moon/distance_earth_moon) * (R_moon / distance_earth_moon) * albedo_moon / (4*Math.PI);
            return FL0;
        }

        // 플랑크 함수
        private double logPlanckBB(double lambda, double temp)
        {
            // lambda : 관측 파장 (μm)
            // temp : 유효온도
            double freqGHz = 2.998e+5 / lambda;
            double beta = (6.626 / 1.381) / 100;
            double logh = Math.Log10(6.626) - 34;
            double logc = Math.Log10(2.998) + 8;
            double logB = Math.Log10(2) + logh - 2 * logc + 3 * Math.Log10(freqGHz) + 27 - Math.Log10(Math.Exp(beta * (freqGHz / temp)) - 1);
            return logB;
        }

        // 달의 반사율
        private double AlbdedoMoon(double alpha, double scattering_albedo = 0.12)
        {
            // scattering_albedo : 달 표면의 산란율 (0.1~0.15 범위에서 조정 가능)
            if (alpha > 180) alpha = 360 - alpha;
            double phase_function = Math.Sin(Deg2Rad(alpha)) - Deg2Rad(alpha) * Math.Cos(Deg2Rad(alpha));
            return scattering_albedo * phase_function;
        }
        private double MoonlitIntensity(double obsZenithDist, double moonZenithDist, double theta, double FL0, double lambda=0.55)
        {
            // lambda : 관측 파장 (μm)
            // FL0 : 대기 상단 달의 flux

            // 달빛 산란 모델 (Rayleigh + Mie)
            double tauR = RayleighOpticalDepth(moonZenithDist, lambda);
            double tauA = AerosolOpticalDepth(moonZenithDist, lambda);
            double P_R = RayleighPhaseFunction(theta);
            double P_M = MiePhaseFunction(theta);

            double P_total = ( tauA*P_M + tauR*P_R )/(tauA + tauR);
            double tau_total = tauA + tauR;

            // 달빛 산란 강도 계산
            double numerator = Math.Exp(-tau_total / Math.Cos(Deg2Rad(obsZenithDist))) - Math.Exp(-tau_total / Math.Cos(Deg2Rad(moonZenithDist)));
            double denominator = 1/Math.Cos(Deg2Rad(moonZenithDist)) - 1/Math.Cos(Deg2Rad(obsZenithDist));
            double I_scattered = FL0 * P_total * (tauA / tau_total) * (1 / Math.Cos(Deg2Rad(obsZenithDist))) * (numerator / denominator);

            //Debug.WriteLine($"num: {numerator}, den: {denominator}, secz: {(1 / Math.Cos(Deg2Rad(obsZenithDist)))}, P_total: {P_total}");
            //Debug.WriteLine($"l_scattered: {I_scattered}");

            return I_scattered;
        }

        // ---------------------- 달 산란 관련 함수 ----------------------
        // Rayleigh 산란 phase function
        private double RayleighPhaseFunction(double theta)
        {
            double cosTheta = Math.Cos(Deg2Rad(theta));
            double chi = 0.0148;

            double coef = 3 * (1 - chi) / (16 * Math.PI * (1 + 2 * chi));
            double result = coef * ((1 + 3 * chi) / (1 - chi) + cosTheta * cosTheta);
            return result;
        }

        // Mie 산란 phase function (Henyey-Greenstein phase function 근사)
        private double MiePhaseFunction(double theta)
        {
            double cosTheta = Math.Cos(Deg2Rad(theta));
            double g = 0.9; // 평균 구름 입자에 대한 g 값
            double numerator = 1 - g * g;
            double denominator = Math.Pow(1 + g * g - 2 * g * cosTheta, 1.5);
            double result = (numerator / (4 * Math.PI * denominator));
            return result;
        }

        // Rayleigh optical depth at λ (μm), pressure P (hPa); then slant τ = m * τ_vert
        // Bucholtz (1995) / Bodhaine et al. (1999) 계열 공식
        private double RayleighOpticalDepth(double zenithDistance, double wavelengthMicron, double pressure_hPa = 1013.25)
        {
            double l = wavelengthMicron;
            double l2 = l * l;
            double term = 1.0 + 0.0113 / l2 + 0.00013 / (l2 * l2);
            double tauR_vert = (pressure_hPa / 1013.25) * 0.008569 * Math.Pow(l, -4.0) * term;
            return tauR_vert;
        }

        // Aerosol (Mie) optical depth via Ångström law; default τ(550nm)=0.1, α=1.3
        private double AerosolOpticalDepth(double zenithDistance, double wavelengthMicron, double tau550 = 0.10, double angstromAlpha = 1.3)
        {
            double tau_vert = tau550 * Math.Pow(wavelengthMicron / 0.55, -angstromAlpha);
            return tau_vert;
        }
        // --------------------------------------------------------


        // 천정거리 계산 메서드
        private double ZenithDistance(double RA, double Dec, DateTime time)
        {
            var site = form.GetObservingSite();
            double latitude = site.lat;
            double cosz = Math.Sin(Deg2Rad(latitude)) * Math.Sin(Deg2Rad(Dec)) +
                          Math.Cos(Deg2Rad(latitude)) * Math.Cos(Deg2Rad(Dec)) * Math.Cos(Deg2Rad(HA_cal(RA, Dec, time)));
            cosz = Math.Min(1.0, Math.Max(-1.0, cosz)); // 수치적 안정성 확보
            return Math.Acos(cosz) * 180.0 / Math.PI; // 도 단위로 반환
        }
        // ----------- 25.08.31~25.09.03 수정(끝) -----------


        // 달의 적경, 적위 계산
        // Meeus "Astronomical Algorithms" 기반
        // 지정한 시간에 대한 달의 적경(시간 단위)과 적위(도 단위)를 반환
        public static (double RA, double Dec) GetMoonEquatorialCoordinates(DateTime utc)
        {
            // 1. 율리우스일(JD) 계산
            double jd = DateTimeToJulianDate(utc);

            // 2. T (J2000.0 기준 유효일수/36525)
            double T = (jd - 2451545.0) / 36525.0;

            // 3. 달의 평균 황경, 평균 이심, 평균 거리, 평균 승교점 경도 (deg)
            double L1 = 218.3164477 + 481267.88123421 * T
                - 0.0015786 * T * T + T * T * T / 538841.0 - T * T * T * T / 65194000.0;
            double D = 297.8501921 + 445267.1114034 * T
                - 0.0018819 * T * T + T * T * T / 545868.0 - T * T * T * T / 113065000.0;
            double M = 357.5291092 + 35999.0502909 * T
                - 0.0001536 * T * T + T * T * T / 24490000.0;
            double M1 = 134.9633964 + 477198.8675055 * T
                + 0.0087414 * T * T + T * T * T / 69699.0 - T * T * T * T / 14712000.0;
            double F = 93.2720950 + 483202.0175233 * T
                - 0.0036539 * T * T - T * T * T / 3526000.0 + T * T * T * T / 863310000.0;

            // 각도를 0~360도로 정규화
            L1 = NormalizeAngle(L1);
            D = NormalizeAngle(D);
            M = NormalizeAngle(M);
            M1 = NormalizeAngle(M1);
            F = NormalizeAngle(F);

            // 4. 달의 황경, 황위(근사, 주요 항만 사용)
            // (정밀 계산은 Meeus 47장, 여기서는 대표 5개 항만 사용)
            double lon = L1
                + 6.289 * Math.Sin(Deg2Rad(M1))
                + 1.274 * Math.Sin(Deg2Rad(2 * D - M1))
                + 0.658 * Math.Sin(Deg2Rad(2 * D))
                + 0.214 * Math.Sin(Deg2Rad(2 * M1))
                - 0.186 * Math.Sin(Deg2Rad(M))
                - 0.059 * Math.Sin(Deg2Rad(2 * D - 2 * M1));

            double lat = 5.128 * Math.Sin(Deg2Rad(F))
                + 0.280 * Math.Sin(Deg2Rad(M1 + F))
                + 0.277 * Math.Sin(Deg2Rad(M1 - F))
                + 0.173 * Math.Sin(Deg2Rad(2 * D - F))
                + 0.055 * Math.Sin(Deg2Rad(2 * D + F - M1));

            // 5. 황도좌표 → 적도좌표 변환 (경사각)
            double eps = 23.439291 - 0.0130042 * T; // J2000.0 기준 평균 황도경사

            // 적위(Dec)
            double sinDec = Math.Sin(Deg2Rad(lat)) * Math.Cos(Deg2Rad(eps))
                + Math.Cos(Deg2Rad(lat)) * Math.Sin(Deg2Rad(eps)) * Math.Sin(Deg2Rad(lon));
            double dec = Math.Asin(sinDec) * 180.0 / Math.PI;

            // 적경(RA)
            double y = Math.Sin(Deg2Rad(lon)) * Math.Cos(Deg2Rad(eps))
                - Math.Tan(Deg2Rad(lat)) * Math.Sin(Deg2Rad(eps));
            double x = Math.Cos(Deg2Rad(lon));
            double ra = Math.Atan2(y, x) * 180.0 / Math.PI;
            if (ra < 0) ra += 360.0;
            ra /= 15.0; // 시간 단위

            return (ra, dec);
        }

        // 태양의 적경, 적위 계산
        public static (double ra, double dec) GetSunEquatorialCoordinates(DateTime utc)
        {
            // 1. 율리우스일(JD) 계산
            double jd = DateTimeToJulianDate(utc);

            // 2. T (J2000.0 기준 유효일수/36525)
            double T = (jd - 2451545.0) / 36525.0;

            // 2) Mean longitude & anomaly (deg)
            double L0 = 280.46646 + 36000.76983 * T + 0.0003032 * T * T;
            double M = 357.52911 + 35999.05029 * T - 0.0001537 * T * T - 0.00000048 * T * T * T;

            // 3) Equation of center (deg)
            double Mr = Deg2Rad(M);
            double C = (1.914602 - 0.004817 * T - 0.000014 * T * T) * Math.Sin(Mr)
                     + (0.019993 - 0.000101 * T) * Math.Sin(2 * Mr)
                     + 0.000289 * Math.Sin(3 * Mr);

            // 4) True & apparent longitude (deg)
            double Theta = L0 + C;
            double Omega = 125.04 - 1934.136 * T;
            double lambda = Theta - 0.00569 - 0.00478 * Math.Sin(Deg2Rad(Omega));
            lambda = NormalizeAngle(lambda);

            // 5) Obliquity (deg)
            double eps0 = 23.0 + 26.0 / 60.0 + 21.448 / 3600.0
                        - (46.8150 / 3600.0) * T - (0.00059 / 3600.0) * T * T + (0.001813 / 3600.0) * T * T * T;
            double eps = eps0 + 0.00256 * Math.Cos(Deg2Rad(Omega));

            // 6) Ecliptic -> Equatorial (β≈0)
            double lam = Deg2Rad(lambda);
            double epsr = Deg2Rad(eps);
            double sinDec = Math.Sin(lam) * Math.Sin(epsr);
            double dec = Rad2Deg(Math.Asin(sinDec));

            double y = Math.Sin(lam) * Math.Cos(epsr);
            double x = Math.Cos(lam);
            double ra = Rad2Deg(Math.Atan2(y, x));
            ra = NormalizeAngle(ra) / 15; // hours

            return (ra, dec);
        }


        // 각도(도) → 라디안
        private static double Deg2Rad(double deg) => deg * Math.PI / 180.0;

        private static double Rad2Deg(double rad) => rad * 180.0 / Math.PI;

        // 0~360도 정규화
        private static double NormalizeAngle(double deg)
        {
            deg %= 360.0;
            if (deg < 0) deg += 360.0;
            return deg;
        }


        // DateTime → Julian Date (UTC 기준)
        private static double DateTimeToJulianDate(DateTime date)
        {
            int Y = date.Year;
            int M = date.Month;
            double D = date.Day + (date.Hour + (date.Minute + date.Second / 60.0) / 60.0) / 24.0;

            if (M <= 2)
            {
                Y -= 1;
                M += 12;
            }
            int A = Y / 100;
            int B = 2 - A + (A / 4);
            double JD = Math.Floor(365.25 * (Y + 4716)) + Math.Floor(30.6001 * (M + 1)) + D + B - 1524.5;
            return JD;
        }


        // 장애물 고도를 선형 보간하여 계산하는 메서드
        private double InterpolateObstacleAltitude(double azimuth, List<double> obstacleAzList, List<double> obstacleAltList)
        {
            if (obstacleAzList == null || obstacleAltList == null || obstacleAzList.Count < 2)
            {
                return 0.0; // 장애물 데이터가 없으면 고도 0 반환
            }

            // 방위각을 0~360도 범위로 정규화
            while (azimuth < 0) azimuth += 360;
            while (azimuth >= 360) azimuth -= 360;

            // 가장 가까운 두 점을 찾아서 선형 보간
            int n = obstacleAzList.Count;

            // 정확히 일치하는 방위각이 있는지 확인
            for (int i = 0; i < n; i++)
            {
                if (Math.Abs(obstacleAzList[i] - azimuth) < 0.001)
                {
                    return obstacleAltList[i];
                }
            }

            // 선형 보간을 위한 두 점 찾기
            int leftIndex = -1;
            int rightIndex = -1;

            for (int i = 0; i < n; i++)
            {
                double currentAz = obstacleAzList[i];

                if (currentAz <= azimuth)
                {
                    if (leftIndex == -1 || obstacleAzList[leftIndex] < currentAz)
                    {
                        leftIndex = i;
                    }
                }

                if (currentAz >= azimuth)
                {
                    if (rightIndex == -1 || obstacleAzList[rightIndex] > currentAz)
                    {
                        rightIndex = i;
                    }
                }
            }

            // 경계 조건 처리 (0도/360도 경계)
            if (leftIndex == -1 || rightIndex == -1)
            {
                // 0도 근처에서 360도와 0도 사이의 보간이 필요한 경우
                double maxAz = obstacleAzList.Max();
                double minAz = obstacleAzList.Min();

                if (azimuth > maxAz)
                {
                    // azimuth가 최대값보다 큰 경우, 최대값과 최소값(+360) 사이 보간
                    leftIndex = obstacleAzList.IndexOf(maxAz);
                    rightIndex = obstacleAzList.IndexOf(minAz);

                    double leftAz = obstacleAzList[leftIndex];
                    double rightAz = obstacleAzList[rightIndex] + 360; // 360도 추가
                    double leftAlt = obstacleAltList[leftIndex];
                    double rightAlt = obstacleAltList[rightIndex];

                    double ratio = (azimuth - leftAz) / (rightAz - leftAz);
                    return leftAlt + ratio * (rightAlt - leftAlt);
                }
                else if (azimuth < minAz)
                {
                    // azimuth가 최소값보다 작은 경우, 최대값(-360)과 최소값 사이 보간
                    leftIndex = obstacleAzList.IndexOf(maxAz);
                    rightIndex = obstacleAzList.IndexOf(minAz);

                    double leftAz = obstacleAzList[leftIndex] - 360; // 360도 빼기
                    double rightAz = obstacleAzList[rightIndex];
                    double leftAlt = obstacleAltList[leftIndex];
                    double rightAlt = obstacleAltList[rightIndex];

                    double ratio = (azimuth - leftAz) / (rightAz - leftAz);
                    return leftAlt + ratio * (rightAlt - leftAlt);
                }
            }

            // 일반적인 선형 보간
            if (leftIndex != -1 && rightIndex != -1 && leftIndex != rightIndex)
            {
                double leftAz = obstacleAzList[leftIndex];
                double rightAz = obstacleAzList[rightIndex];
                double leftAlt = obstacleAltList[leftIndex];
                double rightAlt = obstacleAltList[rightIndex];

                double ratio = (azimuth - leftAz) / (rightAz - leftAz);
                return leftAlt + ratio * (rightAlt - leftAlt);
            }

            // 기본값: 가장 가까운 점의 고도 반환
            double minDistance = double.MaxValue;
            double nearestAltitude = 0.0;

            for (int i = 0; i < n; i++)
            {
                double distance = Math.Min(
                    Math.Abs(obstacleAzList[i] - azimuth),
                    360 - Math.Abs(obstacleAzList[i] - azimuth)
                );

                if (distance < minDistance)
                {
                    minDistance = distance;
                    nearestAltitude = obstacleAltList[i];
                }
            }

            return nearestAltitude;
        }


        private double HA_cal(double RA, double Dec, DateTime time)
        {
            // 원본 RA입력은 시간 단위
            // RA는 시간 단위로 입력되므로, 15를 곱해 도 단위로 변환
            // 라이브러리 사용: AstroAlgo.Basic.CoordinateSystem.GetHourAngle
            // 관측지 경도 (서울 127.0도, 동경 양수)
            // 라이브러리 사용하면 시간각이 -180~180도 범위로 자동 정규화됨
            var site = form.GetObservingSite(); // Form1의 함수 호출
            double ha = CoordinateSystem.GetHourAngle(time, RA * 15, site.lon);

            return ha;
        }


        // 이동하는데 걸리는 시간 계산
        private double MovingTime(double RA1, double Dec1, double RA2, double Dec2, DateTime time)
        {
            // *** 각도 단위 : 도(degree)
            // 기기 사양
            // 기기 사양
            double MotorSpeedRA = 10.0;
            double MotorSpeedDec = 10.0;
            double MotorAccelerationRA = 1.2;
            double MotorAccelerationDec = 1.2;

            // textBox9~12 값이 유효하면 해당 값 사용
            if (form != null)
            {
                double v;
                v = form.MotorSpeedRAValue;
                if (v > 0) MotorSpeedRA = v;
                v = form.MotorSpeedDecValue;
                if (v > 0) MotorSpeedDec = v;
                v = form.MotorAccelerationRAValue;
                if (v > 0) MotorAccelerationRA = v;
                v = form.MotorAccelerationDecValue;
                if (v > 0) MotorAccelerationDec = v;
            }

            // 최대 속도에 도달하기 까지 걸리는 시간
            double t0RA = MotorSpeedRA / MotorAccelerationRA;
            double t0Dec = MotorSpeedDec / MotorAccelerationDec;

            bool meridianpassing = false;
            if (HA_cal(RA1, Dec1, time) * HA_cal(RA2, Dec2, time) < 0) meridianpassing = true;


            // 시간 단위 RA → 도 단위로 변환
            double RA1_deg = RA1 * 15.0;
            double RA2_deg = RA2 * 15.0;

            // 움직인 시간 계산을 위해 적경 범위 변환 : [0, 360) --> [-180, 180)
            // 적경 범위 다시 구현!!!!
            RA1_deg = ((RA1_deg + 180) % 360) - 180;
            RA2_deg = ((RA2_deg + 180) % 360) - 180;


            // 각 모터가움직인 시간들
            double movingtimeRA = 0.0;
            double movingtimeDec = 0.0;
            if (meridianpassing) // 자오선 넘김이 있을 경우
            {
                // 자오선 넘김이 있을 경우
                double RA1_op = ((RA1_deg + 180) % 360) - 180;
                double RA2_op = ((RA2_deg + 180) % 360) - 180;
                if (Math.Abs(RA1_op - RA2_op) < MotorAccelerationRA * t0RA * t0RA)
                    movingtimeRA += 2 * Math.Sqrt(Math.Abs(RA1_op - RA2_op) / MotorAccelerationRA);
                else
                {
                    movingtimeRA += 2 * t0RA;
                    movingtimeRA += (Math.Abs(RA1_op - RA2_op) - MotorAccelerationRA * t0RA * t0RA) / MotorSpeedRA;
                }


                double decZenith2 = 90 * 2.0;
                if ((decZenith2 - Dec1 - Dec2) < MotorAccelerationDec * t0Dec * t0Dec) // 최대속도 도달 전에 다시 감속해야 할 경우
                    movingtimeDec += 2 * Math.Sqrt((decZenith2 - Dec1 - Dec2) / MotorAccelerationDec);
                else
                {
                    movingtimeDec += 2 * t0Dec;
                    movingtimeDec += ((decZenith2 - Dec1 - Dec2) - MotorAccelerationDec * t0Dec * t0Dec) / MotorSpeedDec;
                }
            }
            else // 자오선 넘김이 없을 경우
            {
                // movingtimeRA 계산
                if (Math.Abs(RA1 - RA2) < MotorAccelerationRA * t0RA * t0RA) // 최대속도 도달 전에 다시 감속해야 할 경우
                    movingtimeRA += 2 * Math.Sqrt(Math.Abs(RA1 - RA2) / MotorAccelerationRA);
                else
                {
                    movingtimeRA += 2 * t0RA;
                    movingtimeRA += (Math.Abs(RA1 - RA2) - MotorAccelerationRA * t0RA * t0RA) / MotorSpeedRA;
                }

                // movingtimeDec 계산
                if (Math.Abs(Dec1 - Dec2) < MotorAccelerationDec * t0Dec * t0Dec) // 최대속도 도달 전에 다시 감속해야 할 경우
                    movingtimeDec += 2 * Math.Sqrt(Math.Abs(Dec1 - Dec2) / MotorAccelerationDec);
                else
                {
                    movingtimeDec += 2 * t0Dec;
                    movingtimeDec += (Math.Abs(Dec1 - Dec2) - MotorAccelerationDec * t0Dec * t0Dec) / MotorSpeedDec;
                }
            }

            return Math.Max(movingtimeRA, movingtimeDec);
        }
    }


    // 시간 입력을 위한 새로운 폼 클래스
    public class TimeInputForm : Form
    {
        public DateTime SelectedDateTime { get; private set; }
        private DateTimePicker dateTimePicker;
        private Button okButton;
        private Button cancelButton;

        public TimeInputForm(DateTime? initial = null)
        {
            this.Text = "시간 입력";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Width = 300;
            this.Height = 150;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            dateTimePicker = new DateTimePicker
            {
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "yyyy-MM-dd HH:mm:ss",
                Value = initial ?? DateTime.Now,
                Width = 200,
                Location = new System.Drawing.Point(40, 20),
                ShowUpDown = true
            };
            this.Controls.Add(dateTimePicker);

            okButton = new Button
            {
                Text = "확인",
                DialogResult = DialogResult.OK,
                Location = new System.Drawing.Point(50, 60),
                Width = 80
            };
            okButton.Click += (s, e) => { SelectedDateTime = dateTimePicker.Value; this.DialogResult = DialogResult.OK; };
            this.Controls.Add(okButton);

            cancelButton = new Button
            {
                Text = "취소",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(150, 60),
                Width = 80
            };
            this.Controls.Add(cancelButton);
        }
    }


    // 수동 입력 폼 클래스
    public class ManualTargetInputForm : Form
    {
        public string TargetName { get; private set; }
        public double RA { get; private set; }
        public double Dec { get; private set; }
        public double ObsTime { get; private set; }

        private TextBox nameBox, raBox, decBox, obsTimeBox;
        private Button okButton, cancelButton;

        public ManualTargetInputForm()
        {
            this.Text = "관측 대상 수동 추가";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Width = 320;
            this.Height = 230;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            Label nameLabel = new Label { Text = "대상 이름", Left = 20, Top = 20, Width = 100 };
            nameBox = new TextBox { Left = 130, Top = 18, Width = 170 };

            Label raLabel = new Label { Text = "적경(RA, 시간)", Left = 20, Top = 55, Width = 100 };
            raBox = new TextBox { Left = 130, Top = 53, Width = 170, Text = "0" };

            Label decLabel = new Label { Text = "적위(Dec, 도)", Left = 20, Top = 90, Width = 100 };
            decBox = new TextBox { Left = 130, Top = 88, Width = 170, Text = "0" };

            Label obsTimeLabel = new Label { Text = "관측 시간(분)", Left = 20, Top = 125, Width = 100 };
            obsTimeBox = new TextBox { Left = 130, Top = 123, Width = 170, Text = "10" };

            okButton = new Button { Text = "확인", Left = 60, Top = 160, Width = 80, DialogResult = DialogResult.OK };
            cancelButton = new Button { Text = "취소", Left = 160, Top = 160, Width = 80, DialogResult = DialogResult.Cancel };

            okButton.Click += (s, e) =>
            {
                TargetName = nameBox.Text.Trim();
                double ra, dec, obs;
                double.TryParse(raBox.Text, out ra);
                double.TryParse(decBox.Text, out dec);
                double.TryParse(obsTimeBox.Text, out obs);
                RA = ra;
                Dec = dec;
                ObsTime = obs;
                this.DialogResult = DialogResult.OK;
            };

            this.Controls.AddRange(new Control[] { nameLabel, nameBox, raLabel, raBox, decLabel, decBox, obsTimeLabel, obsTimeBox, okButton, cancelButton });
        }
    }

}


// final update : 0913 00:55 / crane206265