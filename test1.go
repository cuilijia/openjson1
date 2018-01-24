package openjson

import (
	"encoding/json"
	"fmt"
	"github.com/tealeg/xlsx"
	"strconv"
)

type Presentation struct {
	Zh   string  `json:"zh-TW"`
	En   string  `json:"en"`
}

type Score struct {
	N1	string	`json:"1"`
	N2	string	`json:"2"`
	N3	string	`json:"3"`
	N4	string	`json:"4"`
	N5	string	`json:"5"`
	N6	string	`json:"6"`
	N7	string	`json:"7"`
	N8	string	`json:"8"`
	N9	string	`json:"9"`
	N10	string	`json:"10"`
	N11	string	`json:"11"`
	N12	string	`json:"12"`
	N13	string	`json:"13"`
	N14	string	`json:"14"`
	N15	string	`json:"15"`
	N16	string	`json:"16"`
	N17	string	`json:"17"`
	N18	string	`json:"18"`
	N19	string	`json:"19"`
	N20	string	`json:"20"`
	N21	string	`json:"21"`
	N22	string	`json:"22"`
	N23	string	`json:"23"`
	N24	string	`json:"24"`
	N25	string	`json:"25"`
	N26	string	`json:"26"`
	N27	string	`json:"27"`
	N28	string	`json:"28"`
	N29	string	`json:"29"`
	N30	string	`json:"30"`
	N31	string	`json:"31"`
	N32	string	`json:"32"`
	N33	string	`json:"33"`
	N34	string	`json:"34"`
	N35	string	`json:"35"`
	N36	string	`json:"36"`
	N37	string	`json:"37"`
	N38	string	`json:"38"`
	N39	string	`json:"39"`
	N40	string	`json:"40"`
	N41	string	`json:"41"`
	N42	string	`json:"42"`
	N43	string	`json:"43"`
	N44	string	`json:"44"`
	N45	string	`json:"45"`
	N46	string	`json:"46"`
	N47	string	`json:"47"`
	N48	string	`json:"48"`
	N49	string	`json:"49"`
	N50	string	`json:"50"`
	N51	string	`json:"51"`
	N52	string	`json:"52"`
	N53	string	`json:"53"`
	N54	string	`json:"54"`
	N55	string	`json:"55"`
	N56	string	`json:"56"`
	N57	string	`json:"57"`
	N58	string	`json:"58"`
	N59	string	`json:"59"`
	N60	string	`json:"60"`
	N61	string	`json:"61"`
	N62	string	`json:"62"`
	N63	Presentation	`json:"63"`
	N64	Presentation	`json:"64"`
	N65	Presentation	`json:"65"`
	N66	Presentation	`json:"66"`
	N67	string	`json:"67"`
	N68	string	`json:"68"`
	N69	string	`json:"69"`
	N70	string	`json:"70"`
	N71	string	`json:"71"`
	N72	string	`json:"72"`
	N73	string	`json:"73"`
	N74	string	`json:"74"`
	N75	string	`json:"75"`
	N76	string	`json:"76"`
	N77	string	`json:"77"`
	N78	string	`json:"78"`
	N79	string	`json:"79"`
	N80	string	`json:"80"`
	N81	string	`json:"81"`
	N82	string	`json:"82"`
	N83	string	`json:"83"`
	N84	string	`json:"84"`
	N85	string	`json:"85"`
	N86	string	`json:"86"`
	N87	string	`json:"87"`
	N88	string	`json:"88"`
	N89	string	`json:"89"`
	N90	string	`json:"90"`
	N91	string	`json:"91"`
	N92	string	`json:"92"`
	N93	string	`json:"93"`
	N94	string	`json:"94"`
	N95	string	`json:"95"`
	N96	string	`json:"96"`
	N97	string	`json:"97"`
	N98	string	`json:"98"`
	N99	string	`json:"99"`
	N100	string	`json:"100"`
	N101	string	`json:"101"`
	N102	string	`json:"102"`
	N103	string	`json:"103"`
	N104	string	`json:"104"`
	N105	string	`json:"105"`
	N106	string	`json:"106"`
	N107	string	`json:"107"`
	N108	string	`json:"108"`
	N109	string	`json:"109"`
	N110	string	`json:"110"`
	N111	string	`json:"111"`
	N112	string	`json:"112"`
	N113	string	`json:"113"`
	N114	string	`json:"114"`
	N115	string	`json:"115"`
	N116	string	`json:"116"`
	N117	string	`json:"117"`
	N118	string	`json:"118"`
	N119	string	`json:"119"`
	N120	string	`json:"120"`
	N121	string	`json:"121"`
	N122	string	`json:"122"`
	N123	string	`json:"123"`
	N124	string	`json:"124"`
	N125	string	`json:"125"`
	N126	string	`json:"126"`
	N127	string	`json:"127"`
	N128	string	`json:"128"`
	N129	string	`json:"129"`
	N130	string	`json:"130"`
	N131	string	`json:"131"`
	N132	string	`json:"132"`
	N133	string	`json:"133"`
	N134	Presentation	`json:"134"`
	N135	string	`json:"135"`
	N136	Presentation	`json:"136"`
	N137	string	`json:"137"`
	N138	Presentation	`json:"138"`
	N139	Presentation	`json:"139"`
	N140	Presentation	`json:"140"`
	N141	Presentation	`json:"141"`
	N142	Presentation	`json:"142"`
	N143	string	`json:"143"`
	N144	string	`json:"144"`
	N145	string	`json:"145"`
	N146	string	`json:"146"`
	N147	string	`json:"147"`
	N148	string	`json:"148"`
	N149	string	`json:"149"`
	N150	string	`json:"150"`
	N151	string	`json:"151"`
	N152	string	`json:"152"`
	N153	string	`json:"153"`
	N154	string	`json:"154"`
	N155	string	`json:"155"`
	N156	string	`json:"156"`
	N157	string	`json:"157"`
	N158	string	`json:"158"`
	N159	string	`json:"159"`
	N160	string	`json:"160"`
	N161	string	`json:"161"`
	N162	string	`json:"162"`
	N163	string	`json:"163"`
	N164	string	`json:"164"`
	N165	string	`json:"165"`
	N166	string	`json:"166"`
	N167	Presentation	`json:"167"`
	N168	Presentation	`json:"168"`
	N169	Presentation	`json:"169"`
	N170	string	`json:"170"`
	N171	string	`json:"171"`
	N172	Presentation	`json:"172"`
	N173	Presentation	`json:"173"`
	N174	Presentation	`json:"174"`
	N175	string	`json:"175"`
	N176	string	`json:"176"`
	N177	string	`json:"177"`
	N178	string	`json:"178"`
	N179	string	`json:"179"`
	N180	string	`json:"180"`
	N181	string	`json:"181"`
	N182	string	`json:"182"`
	N183	string	`json:"183"`
	N184	string	`json:"184"`
	N185	string	`json:"185"`
	N186	string	`json:"186"`
	N187	string	`json:"187"`
	N188	string	`json:"188"`
	N189	string	`json:"189"`
	N190	string	`json:"190"`
	N191	string	`json:"191"`
	N192	string	`json:"192"`
	N193	string	`json:"193"`
	N194	string	`json:"194"`
	N195	string	`json:"195"`
	N196	Presentation	`json:"196"`
	N197	Presentation	`json:"197"`
	N198	string	`json:"198"`
	N199	string	`json:"199"`
	N200	string	`json:"200"`
	N201	string	`json:"201"`
	N202	string	`json:"202"`
	N203	string	`json:"203"`
	N204	string	`json:"204"`
	N205	string	`json:"205"`
	N206	string	`json:"206"`
	N207	string	`json:"207"`
	N208	string	`json:"208"`
	N209	string	`json:"209"`
	N210	string	`json:"210"`
	N211	string	`json:"211"`
	N212	string	`json:"212"`
	N213	string	`json:"213"`
	N214	string	`json:"214"`
	N215	string	`json:"215"`
	N216	string	`json:"216"`
	N217	string	`json:"217"`
	N218	string	`json:"218"`
	N219	string	`json:"219"`
	N220	string	`json:"220"`
	N221	string	`json:"221"`
	N222	string	`json:"222"`
	N223	string	`json:"223"`
	N224	string	`json:"224"`
	N225	string	`json:"225"`
	N226	string	`json:"226"`
	N227	string	`json:"227"`
	N228	string	`json:"228"`
	N229	string	`json:"229"`
	N230	string	`json:"230"`
	N231	string	`json:"231"`
	N232	string	`json:"232"`
	N233	string	`json:"233"`
	N234	string	`json:"234"`
	N235	string	`json:"235"`
	N236	string	`json:"236"`
	N237	string	`json:"237"`
	N238	string	`json:"238"`
	N239	string	`json:"239"`
	N240	string	`json:"240"`
	N241	string	`json:"241"`
	N242	string	`json:"242"`
	N243	string	`json:"243"`
	N244	string	`json:"244"`
	N245	string	`json:"245"`
	N246	string	`json:"246"`
	N247	string	`json:"247"`
	N248	string	`json:"248"`
	N249	string	`json:"249"`
	N250	string	`json:"250"`
	N251	string	`json:"251"`
	N252	string	`json:"252"`
	N253	string	`json:"253"`
	N254	string	`json:"254"`
	N255	string	`json:"255"`
	N256	string	`json:"256"`
	N257	string	`json:"257"`
	N258	string	`json:"258"`
	N259	string	`json:"259"`
	N260	string	`json:"260"`
	N261	string	`json:"261"`
	N262	string	`json:"262"`
	N263	string	`json:"263"`
	N264	string	`json:"264"`
	N265	string	`json:"265"`
	N266	string	`json:"266"`
	N267	string	`json:"267"`
	N268	string	`json:"268"`
	N269	string	`json:"269"`
	N270	string	`json:"270"`
	N271	string	`json:"271"`
	N272	string	`json:"272"`
	N273	string	`json:"273"`
	N274	string	`json:"274"`
	N275	string	`json:"275"`
	N276	string	`json:"276"`
	N277	string	`json:"277"`
	N278	string	`json:"278"`
	N279	string	`json:"279"`
	N280	string	`json:"280"`
	N281	string	`json:"281"`
	N282	string	`json:"282"`
	N283	string	`json:"283"`
	N284	string	`json:"284"`
	N285	string	`json:"285"`
	N286	string	`json:"286"`
	N287	string	`json:"287"`
	N288	string	`json:"288"`
	N289	string	`json:"289"`
	N290	Presentation	`json:"290"`
	N291	Presentation	`json:"291"`
	N292	string	`json:"292"`
	N293	string	`json:"293"`
	N294	string	`json:"294"`
	N295	string	`json:"295"`
	N296	string	`json:"296"`
	N297	string	`json:"297"`
	N298	string	`json:"298"`
	N299	string	`json:"299"`
	N300	string	`json:"300"`
	N301	string	`json:"301"`
}

type InsuranceCompany struct {
	Id int  `json:"id"`
	Name Presentation   `json:"name"`
	LogoImageUrl string  `json:"Logo_image_url"`
	ProductCategoryIds []int  `json:"Product_category_ids"`
}

type ProductCategory struct {
	Id int  `json:"id"`
	Name Presentation  `json:"name"`
	IconName string  `json:"icon_name"`
}

type Product struct {
	Id int  `json:"id"`
	Name Presentation  `json:"name"`
	ShortNote Presentation  `json:"Short_note"`
	PaymentTerms Presentation  `json:"Payment_terms"`
	IssueAge string  `json:"Issue_age"`
	CurrencyCode string  `json:"Currency_code"`
	ProductWebsite string  `json:"Product_website"`
    Scores Score  `json:"Scores"`
    Company InsuranceCompany  `json:"insurance_company"`
	Category ProductCategory  `json:"product_category"`
}

type MyProduct struct {
	Id int
	NameZHCN string
	NameEN string
	TypeName string
	CompanyId int
	Desc MyProductDetail
}

type MyProductDetail struct {
	ShortNote Presentation
	//PaymentTerms Presentation
	IssueAge string
	Aa string
	//CurrencyCode string
	//ProductWebsite string
	//Scores Score
	//Company InsuranceCompany
	//Category ProductCategory
}

type Productslice struct {
	Products []Product
}

func Princhuxuweijiscore(scorez Score)  {
	//fmt.Println(scorez)
	//fmt.Println("+++++",scorez.N143,"+++++")
	fmt.Println("详细资料：{  赔偿额为已付保费的倍数")
	fmt.Println("身故赔偿(排名：",scorez.N6,"分数：",scorez.N5,"/10.0)： 第1年(",scorez.N1,"*)，第10年(",scorez.N2,"*)，第20年(",scorez.N3,"*)，第30年(",scorez.N4,"*)")
	fmt.Println("保证回报(排名：",scorez.N18,"分数：",scorez.N14,"/10.0)： 保证回本年(",scorez.N7,")，保证年化汇报率{10年(",scorez.N8,"%),20年（",scorez.N9,"%),30年(",scorez.N10,"%)}")
	fmt.Println("预期回报(排名：",scorez.N30,"分数：",scorez.N26,"/10.0)： 预期回本年(",scorez.N19,")，预期年化汇报率{10年(",scorez.N20,"%),20年（",scorez.N21,"%),30年(",scorez.N22,"%)}")
	fmt.Println("危机保障(排名：",scorez.N40,"分数：",scorez.N35,"/10.0)： {")
	fmt.Println("   <1>癌症评分(排名：",scorez.N41,"分数：",scorez.N36,"/10.0)： 早期癌症赔额(",scorez.N49,"*)，第一次癌症赔偿额(",scorez.N48,"*),第二次癌症赔偿额（",scorez.N50,"*),癌症保障条款评分(",scorez.N51,"/10.0),癌症短评:",scorez.N63)
	fmt.Println("   <2>心脏病评分(排名：",scorez.N42,"分数：",scorez.N37,"/10.0)： 早期心脏疾病治疗赔额(",scorez.N53,"*)，手术赔偿后心脏病赔偿(",scorez.N143,"*),第一次心脏病赔偿额（",scorez.N52,"*),第二次心脏病赔偿额(",scorez.N54,"*),心脏病保障条款评分(",scorez.N55,"/10.0),心脏病短评:",scorez.N64)
	fmt.Println("   <3>中风评分(排名：",scorez.N43,"分数：",scorez.N38,"/10.0)： 早期中风赔额(",scorez.N57,"*)，第一次中风赔偿额(",scorez.N56,"*),第二次中风赔偿额（",scorez.N58,"*),中风保障条款评分(",scorez.N59,"/10.0),中风短评:",scorez.N65)
	fmt.Println("   <4>多次严重疾病评分(排名：",scorez.N44,"分数：",scorez.N39,"/10.0)： 癌症+心脏病(",scorez.N147,"*)，癌症+中风(",scorez.N148,"*),癌症+其他主要危疾（",scorez.N150,"*),心脏病+主要中风(",scorez.N149,"*),心脏病+其他主要危疾(",scorez.N151,"*),中风+其他主要危疾(",scorez.N152,"*),其他主要危疾+其他主要危疾(",scorez.N153,"*),短评:",scorez.N66)
	fmt.Println("    }")
	fmt.Println("}")
}

func Princhuxurenshouscore(scorez Score)  {
	//fmt.Println(scorez)
	//fmt.Println("+++++",scorez.N84,"+++++")
	fmt.Println("详细资料：{")
	fmt.Println("身故赔偿(排名：",scorez.N72,"分数：",scorez.N71,"/10.0)： 第1年(",scorez.N67,"*)，第10年(",scorez.N68,"*)，第20年(",scorez.N69,"*)，第30年(",scorez.N70,"*)")
	fmt.Println("保证回报(排名：",scorez.N84,"分数：",scorez.N80,"/10.0)： 保证回本年(",scorez.N73,")，保证年化汇报率{10年(",scorez.N74,"%),20年（",scorez.N75,"%),30年(",scorez.N76,"%)}")
	fmt.Println("预期回报(排名：",scorez.N96,"分数：",scorez.N92,"/10.0)： 预期回本年(",scorez.N85,")，预期年化汇报率{10年(",scorez.N86,"%),20年（",scorez.N87,"%),30年(",scorez.N88,"%)}")
	fmt.Println("}")

}

func Prinpersonalscore(scorez Score)  {
	//fmt.Println(scorez)
	//fmt.Println("+++++",scorez.N134,"+++++")
	fmt.Println("详细资料：{")
	fmt.Println("意外保障(排名：",scorez.N101,"分数：",scorez.N107,"/10.0)： 每年保费（港币）：$",scorez.N113)
	fmt.Println("意外身故/伤残保障(排名：",scorez.N102,"分数：",scorez.N108,"/10.0)： 意外身故/最高伤残赔偿：$",scorez.N114,"；最高赔偿额（年保费倍数）：",scorez.N117,"*")
	fmt.Println("意外受伤医疗保障(排名：",scorez.N103,"分数：",scorez.N109,"/10.0)： 意外受伤医疗保障额：$",scorez.N116,"；最高赔偿额（年保费倍数）：",scorez.N118,"*")
	fmt.Println("意外物理治疗/跌打保障(排名：",scorez.N130,"分数：",scorez.N133,"/10.0)：物理资料保障：",scorez.N134.Zh,"；物理资料赔偿率：",scorez.N135,"%；跌打保障：",scorez.N136.Zh,"；跌打赔偿率：",scorez.N137,"%")
	fmt.Println("意外住院现金(排名：",scorez.N104,"分数：",scorez.N110,"/10.0)：每日住院现金：",scorez.N115,"；最高赔偿率：",scorez.N119,"*")
	fmt.Println("受保运动范围(排名：",scorez.N131,"分数：",scorez.N132,"/10.0)：滑雪：",scorez.N138.Zh,"；潜水：",scorez.N139.Zh,"；跳伞：",scorez.N140.Zh,"；攀石：",scorez.N141.Zh,"；单车竞赛：",scorez.N142.Zh)
	fmt.Println("}")

}

func SpecialPrint1(product Product)  {
	fmt.Println(	"产品[",product.Id,"]号")
	fmt.Println(	"中文名字：",product.Name.Zh)
	fmt.Println(	"英文名字：",product.Name.En)
	fmt.Println(	"产品链接：",product.ProductWebsite)
	fmt.Println(	"公司：",product.Company.Name.Zh,"(中文)/",product.Company.Name.En,"(英文)")
	fmt.Println(	"种类：",product.Category.Name.Zh,"(中文)/",product.Category.Name.En,"(英文)")
	fmt.Println(	"限制年龄：",product.IssueAge)
	fmt.Println(	"支付方式：",product.PaymentTerms.Zh,"(中文)/",product.PaymentTerms.En,"(英文)")
	fmt.Println(	"条件说明：",product.ShortNote.Zh,"(中文)/",product.ShortNote.En,"(英文)")
	fmt.Println(	"暂时码：",product.CurrencyCode)
	fmt.Println("-----------------------------------------------------------")
	fmt.Println()
	file, err := xlsx.OpenFile("test1.xlsx")
	if err != nil {
		panic(err)
	}
	first := file.Sheets[0]
	row := first.AddRow()
	row.SetHeightCM(0.7)
	cell := row.AddCell()
	cell.Value = strconv.Itoa(product.Id)
	cell = row.AddCell()
	cell.Value = product.Name.Zh
	cell = row.AddCell()
	cell.Value = product.Name.En
	cell = row.AddCell()
	cell.Value = product.ProductWebsite
	cell = row.AddCell()
	cell.Value = product.Company.Name.Zh
	cell = row.AddCell()
	cell.Value = product.Company.Name.En
	cell = row.AddCell()
	cell.Value = product.Category.Name.Zh
	cell = row.AddCell()
	cell.Value = product.Category.Name.En
	cell = row.AddCell()
	cell.Value = product.IssueAge
	cell = row.AddCell()
	cell.Value = product.PaymentTerms.Zh
	cell = row.AddCell()
	cell.Value = product.PaymentTerms.En
	cell = row.AddCell()
	cell.Value = product.ShortNote.Zh
	cell = row.AddCell()
	cell.Value = product.ShortNote.En
	cell = row.AddCell()
	cell.Value = product.CurrencyCode

	err = file.Save("test1.xlsx")
	if err != nil {
		panic(err)
	}
}

func Prinmedicalscore(scorez Score)  {
	//fmt.Println(scorez)
	//fmt.Println("+++++",scorez.N184,"+++++")
	fmt.Println("详细资料：{  ")
	fmt.Println("每年赔偿上限(排名：",scorez.N202,"分数：",scorez.N176,"/10.0)： ",scorez.N291.Zh)
	fmt.Println("非严重个案的保障(排名：",scorez.N200,"分数：",scorez.N177,"/10.0)： 上呼吸道感染(",scorez.N184,")，割痔疮手术(",scorez.N185,")，一般受伤(",scorez.N186,")，肠镜或胃镜检查（日常检查）(",scorez.N187,")，肠镜或胃镜检查（入院）(",scorez.N188,")")
	fmt.Println("严重个案的保障(排名：",scorez.N201,"分数：",scorez.N178,"/10.0)： 癌症（原位癌）(",scorez.N189,")，通波仔(",scorez.N190,")，心脏病（严重）(",scorez.N191,")，中风(",scorez.N192,")，严重受伤（入院）(",scorez.N193,")")
	//fmt.Println("预期回报(排名：",scorez.N30,"分数：",scorez.N26,"/10.0)： 预期回本年(",scorez.N19,")，预期年化汇报率{10年(",scorez.N20,"%),20年（",scorez.N21,"%),30年(",scorez.N22,"%)}")
	fmt.Println("平均保费：（0-74） {",scorez.N207,scorez.N217,scorez.N227,scorez.N237,scorez.N247,scorez.N257,scorez.N267,scorez.N277,"}")
    fmt.Println("}")
}

func SpecialPrint2(product Product)  {
	fmt.Println(	"产品[",product.Id,"]号")
	fmt.Println(	"中文名字：",product.Name.Zh)
	fmt.Println(	"英文名字：",product.Name.En)
	fmt.Println(	"产品链接：",product.ProductWebsite)
	fmt.Println(	"公司：",product.Company.Name.Zh,"(中文)/",product.Company.Name.En,"(英文)")
	fmt.Println(	"种类：",product.Category.Name.Zh,"(中文)/",product.Category.Name.En,"(英文)")
	fmt.Println(	"限制年龄：",product.IssueAge)
	fmt.Println(	"支付方式：",product.PaymentTerms.Zh,"(中文)/",product.PaymentTerms.En,"(英文)")
	fmt.Println(	"条件说明：",product.ShortNote.Zh,"(中文)/",product.ShortNote.En,"(英文)")
	fmt.Println(	"暂时码：",product.CurrencyCode)
	//Prinpersonalscore(product.Scores)
	Prinpersonalscore(product.Scores)
	fmt.Println("-----------------------------------------------------------")
	fmt.Println()
}

func writeinpersonalscore(product Product){
	fmt.Println(	"产品[",product.Id,"]号")
	fmt.Println(	"中文名字：",product.Name.Zh)
	fmt.Println(	"英文名字：",product.Name.En)
	fmt.Println(	"产品链接：",product.ProductWebsite)
	fmt.Println(	"公司：",product.Company.Name.Zh,"(中文)/",product.Company.Name.En,"(英文)")
	fmt.Println(	"种类：",product.Category.Name.Zh,"(中文)/",product.Category.Name.En,"(英文)")
	fmt.Println(	"限制年龄：",product.IssueAge)
	fmt.Println(	"支付方式：",product.PaymentTerms.Zh,"(中文)/",product.PaymentTerms.En,"(英文)")
	fmt.Println(	"条件说明：",product.ShortNote.Zh,"(中文)/",product.ShortNote.En,"(英文)")
	fmt.Println(	"暂时码：",product.CurrencyCode)
	var scorez=product.Scores
	fmt.Println("详细资料：{")
	fmt.Println("意外保障(排名：",scorez.N101,"分数：",scorez.N107,"/10.0)： 每年保费（港币）：$",scorez.N113)
	fmt.Println("意外身故/伤残保障(排名：",scorez.N102,"分数：",scorez.N108,"/10.0)： 意外身故/最高伤残赔偿：$",scorez.N114,"；最高赔偿额（年保费倍数）：",scorez.N117,"*")
	fmt.Println("意外受伤医疗保障(排名：",scorez.N103,"分数：",scorez.N109,"/10.0)： 意外受伤医疗保障额：$",scorez.N116,"；最高赔偿额（年保费倍数）：",scorez.N118,"*")
	fmt.Println("意外物理治疗/跌打保障(排名：",scorez.N130,"分数：",scorez.N133,"/10.0)：物理治疗保障：",scorez.N134.Zh,"；物理治疗赔偿率：",scorez.N135,"%；跌打保障：",scorez.N136.Zh,"；跌打赔偿率：",scorez.N137,"%")
	fmt.Println("意外住院现金(排名：",scorez.N104,"分数：",scorez.N110,"/10.0)：每日住院现金：",scorez.N115,"；最高赔偿率：",scorez.N119,"*")
	fmt.Println("受保运动范围(排名：",scorez.N131,"分数：",scorez.N132,"/10.0)：滑雪：",scorez.N138.Zh,"；潜水：",scorez.N139.Zh,"；跳伞：",scorez.N140.Zh,"；攀石：",scorez.N141.Zh,"；单车竞赛：",scorez.N142.Zh)
	fmt.Println("}")
	fmt.Println("---------------------------------------------------------------------------------------------")

	file, err := xlsx.OpenFile("个人意外.xlsx")
	if err != nil {
		panic(err)
	}
	first := file.Sheets[0]
	row := first.AddRow()
	row.SetHeightCM(0.7)
	cell := row.AddCell()
	cell.Value = strconv.Itoa(product.Id)
	cell = row.AddCell()
	cell.Value = product.Name.Zh
	cell = row.AddCell()
	cell.Value = product.Name.En
	cell = row.AddCell()
	cell.Value = product.ProductWebsite
	cell = row.AddCell()
	cell.Value = product.Company.Name.Zh
	cell = row.AddCell()
	cell.Value = product.Company.Name.En
	cell = row.AddCell()
	cell.Value = product.Category.Name.Zh
	cell = row.AddCell()
	cell.Value = product.Category.Name.En
	cell = row.AddCell()
	cell.Value = product.IssueAge
	cell = row.AddCell()
	cell.Value = product.PaymentTerms.Zh
	cell = row.AddCell()
	cell.Value = product.PaymentTerms.En
	cell = row.AddCell()
	cell.Value = product.ShortNote.Zh
	cell = row.AddCell()
	cell.Value = product.ShortNote.En
	cell = row.AddCell()
	cell.Value = product.CurrencyCode
	cell = row.AddCell()
	cell.Value = "详细信息:"
	cell = row.AddCell()
	cell.Value = "意外保障:"
	cell = row.AddCell()
	cell.Value = scorez.N101
	cell = row.AddCell()
	cell.Value = scorez.N107
	cell = row.AddCell()
	cell.Value = scorez.N113
	cell = row.AddCell()
	cell.Value = "意外身故/伤残保障:"
	cell = row.AddCell()
	cell.Value = scorez.N102
	cell = row.AddCell()
	cell.Value = scorez.N108
	cell = row.AddCell()
	cell.Value = scorez.N114
	cell = row.AddCell()
	cell.Value = scorez.N117
	cell = row.AddCell()
	cell.Value = "意外受伤医疗保障:"
	cell = row.AddCell()
	cell.Value = scorez.N103
	cell = row.AddCell()
	cell.Value = scorez.N109
	cell = row.AddCell()
	cell.Value = scorez.N116
	cell = row.AddCell()
	cell.Value = scorez.N118
	cell = row.AddCell()
	cell.Value = "意外物理治疗/跌打保障:"
	cell = row.AddCell()
	cell.Value = scorez.N130
	cell = row.AddCell()
	cell.Value = scorez.N133
	cell = row.AddCell()
	cell.Value = scorez.N134.Zh
	cell = row.AddCell()
	cell.Value = scorez.N135
	cell = row.AddCell()
	cell.Value = scorez.N136.Zh
	cell = row.AddCell()
	cell.Value = scorez.N137
	cell = row.AddCell()
	cell.Value = "意外住院现金:"
	cell = row.AddCell()
	cell.Value = scorez.N104
	cell = row.AddCell()
	cell.Value = scorez.N110
	cell = row.AddCell()
	cell.Value = scorez.N115
	cell = row.AddCell()
	cell.Value = scorez.N119
	cell = row.AddCell()
	cell.Value = "受保运动范围:"
	cell = row.AddCell()
	cell.Value = scorez.N131
	cell = row.AddCell()
	cell.Value = scorez.N132
	cell = row.AddCell()
	cell.Value = scorez.N138.Zh
	cell = row.AddCell()
	cell.Value = scorez.N139.Zh
	cell = row.AddCell()
	cell.Value = scorez.N140.Zh
	cell = row.AddCell()
	cell.Value = scorez.N141.Zh
	cell = row.AddCell()
	cell.Value = scorez.N142.Zh

	err = file.Save("个人意外.xlsx")
	if err != nil {
		panic(err)
	}
}

func writeinchuxurenshouscore(product Product){
	fmt.Println(	"产品[",product.Id,"]号")
	fmt.Println(	"中文名字：",product.Name.Zh)
	fmt.Println(	"英文名字：",product.Name.En)
	fmt.Println(	"产品链接：",product.ProductWebsite)
	fmt.Println(	"公司：",product.Company.Name.Zh,"(中文)/",product.Company.Name.En,"(英文)")
	fmt.Println(	"种类：",product.Category.Name.Zh,"(中文)/",product.Category.Name.En,"(英文)")
	fmt.Println(	"限制年龄：",product.IssueAge)
	fmt.Println(	"支付方式：",product.PaymentTerms.Zh,"(中文)/",product.PaymentTerms.En,"(英文)")
	fmt.Println(	"条件说明：",product.ShortNote.Zh,"(中文)/",product.ShortNote.En,"(英文)")
	fmt.Println(	"暂时码：",product.CurrencyCode)
	var scorez=product.Scores
	fmt.Println("详细资料：{")
	fmt.Println("身故赔偿(排名：",scorez.N72,"分数：",scorez.N71,"/10.0)： 第1年(",scorez.N67,"*)，第10年(",scorez.N68,"*)，第20年(",scorez.N69,"*)，第30年(",scorez.N70,"*)")
	fmt.Println("保证回报(排名：",scorez.N84,"分数：",scorez.N80,"/10.0)： 保证回本年(",scorez.N73,")，保证年化回报率{10年(",scorez.N74,"%),20年（",scorez.N75,"%),30年(",scorez.N76,"%)}")
	fmt.Println("预期回报(排名：",scorez.N96,"分数：",scorez.N92,"/10.0)： 预期回本年(",scorez.N85,")，预期年化回报率{10年(",scorez.N86,"%),20年（",scorez.N87,"%),30年(",scorez.N88,"%)}")
	fmt.Println("}")
	fmt.Println("---------------------------------------------------------------------------------------------")

	file, err := xlsx.OpenFile("储蓄型人寿.xlsx")
	if err != nil {
		panic(err)
	}
	first := file.Sheets[0]
	row := first.AddRow()
	row.SetHeightCM(0.7)
	cell := row.AddCell()
	cell.Value = strconv.Itoa(product.Id)
	cell = row.AddCell()
	cell.Value = product.Name.Zh
	cell = row.AddCell()
	cell.Value = product.Name.En
	cell = row.AddCell()
	cell.Value = product.ProductWebsite
	cell = row.AddCell()
	cell.Value = product.Company.Name.Zh
	cell = row.AddCell()
	cell.Value = product.Company.Name.En
	cell = row.AddCell()
	cell.Value = product.Category.Name.Zh
	cell = row.AddCell()
	cell.Value = product.Category.Name.En
	cell = row.AddCell()
	cell.Value = product.IssueAge
	cell = row.AddCell()
	cell.Value = product.PaymentTerms.Zh
	cell = row.AddCell()
	cell.Value = product.PaymentTerms.En
	cell = row.AddCell()
	cell.Value = product.ShortNote.Zh
	cell = row.AddCell()
	cell.Value = product.ShortNote.En
	cell = row.AddCell()
	cell.Value = product.CurrencyCode
	cell = row.AddCell()
	cell.Value = "详细信息:"
	cell = row.AddCell()
	cell.Value = "身故赔偿:"
	cell = row.AddCell()
	cell.Value = scorez.N72
	cell = row.AddCell()
	cell.Value = scorez.N71
	cell = row.AddCell()
	cell.Value = scorez.N67
	cell = row.AddCell()
	cell.Value = scorez.N68
	cell = row.AddCell()
	cell.Value = scorez.N69
	cell = row.AddCell()
	cell.Value = scorez.N70
	cell = row.AddCell()
	cell.Value = "保证回报:"
	cell = row.AddCell()
	cell.Value = scorez.N84
	cell = row.AddCell()
	cell.Value = scorez.N80
	cell = row.AddCell()
	cell.Value = scorez.N73
	cell = row.AddCell()
	cell.Value = scorez.N74
	cell = row.AddCell()
	cell.Value = scorez.N75
	cell = row.AddCell()
	cell.Value = scorez.N76
	cell = row.AddCell()
	cell.Value = "预期回报:"
	cell = row.AddCell()
	cell.Value = scorez.N96
	cell = row.AddCell()
	cell.Value = scorez.N92
	cell = row.AddCell()
	cell.Value = scorez.N85
	cell = row.AddCell()
	cell.Value = scorez.N86
	cell = row.AddCell()
	cell.Value = scorez.N87
	cell = row.AddCell()
	cell.Value = scorez.N88
	cell = row.AddCell()

	err = file.Save("储蓄型人寿.xlsx")
	if err != nil {
		panic(err)
	}
}

func writeinmedicalscore(product Product){
	fmt.Println(	"产品[",product.Id,"]号")
	fmt.Println(	"中文名字：",product.Name.Zh)
	fmt.Println(	"英文名字：",product.Name.En)
	fmt.Println(	"产品链接：",product.ProductWebsite)
	fmt.Println(	"公司：",product.Company.Name.Zh,"(中文)/",product.Company.Name.En,"(英文)")
	fmt.Println(	"种类：",product.Category.Name.Zh,"(中文)/",product.Category.Name.En,"(英文)")
	fmt.Println(	"限制年龄：",product.IssueAge)
	fmt.Println(	"支付方式：",product.PaymentTerms.Zh,"(中文)/",product.PaymentTerms.En,"(英文)")
	fmt.Println(	"条件说明：",product.ShortNote.Zh,"(中文)/",product.ShortNote.En,"(英文)")
	fmt.Println(	"暂时码：",product.CurrencyCode)
	var scorez=product.Scores
	fmt.Println("详细资料：{  ")
	fmt.Println("每年赔偿上限(排名：",scorez.N202,"分数：",scorez.N176,"/10.0)： ",scorez.N291.Zh)
	fmt.Println("非严重个案的保障(排名：",scorez.N200,"分数：",scorez.N177,"/10.0)： 上呼吸道感染(",scorez.N184,")，割痔疮手术(",scorez.N185,")，一般受伤(",scorez.N186,")，肠镜或胃镜检查（日常检查）(",scorez.N187,")，肠镜或胃镜检查（入院）(",scorez.N188,")")
	fmt.Println("严重个案的保障(排名：",scorez.N201,"分数：",scorez.N178,"/10.0)： 癌症（原位癌）(",scorez.N189,")，通波仔(",scorez.N190,")，心脏病（严重）(",scorez.N191,")，中风(",scorez.N192,")，严重受伤（入院）(",scorez.N193,")")
	fmt.Println("平均保费：（0-74） {",scorez.N207,scorez.N217,scorez.N227,scorez.N237,scorez.N247,scorez.N257,scorez.N267,scorez.N277,"}")
	fmt.Println("}")
	fmt.Println("---------------------------------------------------------------------------------------------")

	file, err := xlsx.OpenFile("高端医疗 亚洲.xlsx")
	if err != nil {
		panic(err)
	}
	first := file.Sheets[0]
	row := first.AddRow()
	row.SetHeightCM(0.7)
	cell := row.AddCell()
	cell.Value = strconv.Itoa(product.Id)
	cell = row.AddCell()
	cell.Value = product.Name.Zh
	cell = row.AddCell()
	cell.Value = product.Name.En
	cell = row.AddCell()
	cell.Value = product.ProductWebsite
	cell = row.AddCell()
	cell.Value = product.Company.Name.Zh
	cell = row.AddCell()
	cell.Value = product.Company.Name.En
	cell = row.AddCell()
	cell.Value = product.Category.Name.Zh
	cell = row.AddCell()
	cell.Value = product.Category.Name.En
	cell = row.AddCell()
	cell.Value = product.IssueAge
	cell = row.AddCell()
	cell.Value = product.PaymentTerms.Zh
	cell = row.AddCell()
	cell.Value = product.PaymentTerms.En
	cell = row.AddCell()
	cell.Value = product.ShortNote.Zh
	cell = row.AddCell()
	cell.Value = product.ShortNote.En
	cell = row.AddCell()
	cell.Value = product.CurrencyCode
	cell = row.AddCell()
	cell.Value = "详细信息:"
	cell = row.AddCell()
	cell.Value = "每年赔偿上限:"
	cell = row.AddCell()
	cell.Value = scorez.N202
	cell = row.AddCell()
	cell.Value = scorez.N176
	cell = row.AddCell()
	cell.Value = scorez.N291.Zh
	cell = row.AddCell()
	cell.Value = scorez.N291.En
	cell = row.AddCell()
	cell.Value = "非严重个案的保障:"
	cell = row.AddCell()
	cell.Value = scorez.N200
	cell = row.AddCell()
	cell.Value = scorez.N177
	cell = row.AddCell()
	cell.Value = scorez.N184
	cell = row.AddCell()
	cell.Value = scorez.N185
	cell = row.AddCell()
	cell.Value = scorez.N186
	cell = row.AddCell()
	cell.Value = scorez.N187
	cell = row.AddCell()
	cell.Value = scorez.N188
	cell = row.AddCell()
	cell.Value = "严重个案的保障"
	cell = row.AddCell()
	cell.Value = scorez.N201
	cell = row.AddCell()
	cell.Value = scorez.N178
	cell = row.AddCell()
	cell.Value = scorez.N189
	cell = row.AddCell()
	cell.Value = scorez.N190
	cell = row.AddCell()
	cell.Value = scorez.N191
	cell = row.AddCell()
	cell.Value = scorez.N192
	cell = row.AddCell()
	cell.Value = scorez.N193
	cell = row.AddCell()
	cell.Value = "平均保费"
	cell = row.AddCell()
	cell.Value = scorez.N207
	cell = row.AddCell()
	cell.Value = scorez.N217
	cell = row.AddCell()
	cell.Value = scorez.N227
	cell = row.AddCell()
	cell.Value = scorez.N237
	cell = row.AddCell()
	cell.Value = scorez.N247
	cell = row.AddCell()
	cell.Value = scorez.N257
	cell = row.AddCell()
	cell.Value = scorez.N267
	cell = row.AddCell()
	cell.Value = scorez.N277

	err = file.Save("高端医疗 亚洲.xlsx")
	if err != nil {
		panic(err)
	}
}

func writeinchuxuweijiscore(product Product){
	fmt.Println(	"产品[",product.Id,"]号")
	fmt.Println(	"中文名字：",product.Name.Zh)
	fmt.Println(	"英文名字：",product.Name.En)
	fmt.Println(	"产品链接：",product.ProductWebsite)
	fmt.Println(	"公司：",product.Company.Name.Zh,"(中文)/",product.Company.Name.En,"(英文)")
	fmt.Println(	"种类：",product.Category.Name.Zh,"(中文)/",product.Category.Name.En,"(英文)")
	fmt.Println(	"限制年龄：",product.IssueAge)
	fmt.Println(	"支付方式：",product.PaymentTerms.Zh,"(中文)/",product.PaymentTerms.En,"(英文)")
	fmt.Println(	"条件说明：",product.ShortNote.Zh,"(中文)/",product.ShortNote.En,"(英文)")
	fmt.Println(	"暂时码：",product.CurrencyCode)
	var scorez=product.Scores
	fmt.Println("详细资料：{  赔偿额为已付保费的倍数")
	fmt.Println("身故赔偿(排名：",scorez.N6,"分数：",scorez.N5,"/10.0)： 第1年(",scorez.N1,"*)，第10年(",scorez.N2,"*)，第20年(",scorez.N3,"*)，第30年(",scorez.N4,"*)")
	fmt.Println("保证回报(排名：",scorez.N18,"分数：",scorez.N14,"/10.0)： 保证回本年(",scorez.N7,")，保证年化回报率{10年(",scorez.N8,"%),20年（",scorez.N9,"%),30年(",scorez.N10,"%)}")
	fmt.Println("预期回报(排名：",scorez.N30,"分数：",scorez.N26,"/10.0)： 预期回本年(",scorez.N19,")，预期年化回报率{10年(",scorez.N20,"%),20年（",scorez.N21,"%),30年(",scorez.N22,"%)}")
	fmt.Println("危机保障(排名：",scorez.N40,"分数：",scorez.N35,"/10.0)： {")
	fmt.Println("   <1>癌症评分(排名：",scorez.N41,"分数：",scorez.N36,"/10.0)： 早期癌症赔额(",scorez.N49,"*)，第一次癌症赔偿额(",scorez.N48,"*),第二次癌症赔偿额（",scorez.N50,"*),癌症保障条款评分(",scorez.N51,"/10.0),癌症短评:",scorez.N63)
	fmt.Println("   <2>心脏病评分(排名：",scorez.N42,"分数：",scorez.N37,"/10.0)： 早期心脏疾病治疗赔额(",scorez.N53,"*)，手术赔偿后心脏病赔偿(",scorez.N143,"*),第一次心脏病赔偿额（",scorez.N52,"*),第二次心脏病赔偿额(",scorez.N54,"*),心脏病保障条款评分(",scorez.N55,"/10.0),心脏病短评:",scorez.N64)
	fmt.Println("   <3>中风评分(排名：",scorez.N43,"分数：",scorez.N38,"/10.0)： 早期中风赔额(",scorez.N57,"*)，第一次中风赔偿额(",scorez.N56,"*),第二次中风赔偿额（",scorez.N58,"*),中风保障条款评分(",scorez.N59,"/10.0),中风短评:",scorez.N65)
	fmt.Println("   <4>多次严重疾病评分(排名：",scorez.N44,"分数：",scorez.N39,"/10.0)： 癌症+心脏病(",scorez.N147,"*)，癌症+中风(",scorez.N148,"*),癌症+其他主要危疾（",scorez.N150,"*),心脏病+主要中风(",scorez.N149,"*),心脏病+其他主要危疾(",scorez.N151,"*),中风+其他主要危疾(",scorez.N152,"*),其他主要危疾+其他主要危疾(",scorez.N153,"*),短评:",scorez.N66)
	fmt.Println("    }")
	fmt.Println("}")
	fmt.Println("---------------------------------------------------------------------------------------------")

	file, err := xlsx.OpenFile("储蓄型危疾.xlsx")
	if err != nil {
		panic(err)
	}
	first := file.Sheets[0]
	row := first.AddRow()
	row.SetHeightCM(0.7)
	cell := row.AddCell()
	cell.Value = strconv.Itoa(product.Id)
	cell = row.AddCell()
	cell.Value = product.Name.Zh
	cell = row.AddCell()
	cell.Value = product.Name.En
	cell = row.AddCell()
	cell.Value = product.ProductWebsite
	cell = row.AddCell()
	cell.Value = product.Company.Name.Zh
	cell = row.AddCell()
	cell.Value = product.Company.Name.En
	cell = row.AddCell()
	cell.Value = product.Category.Name.Zh
	cell = row.AddCell()
	cell.Value = product.Category.Name.En
	cell = row.AddCell()
	cell.Value = product.IssueAge
	cell = row.AddCell()
	cell.Value = product.PaymentTerms.Zh
	cell = row.AddCell()
	cell.Value = product.PaymentTerms.En
	cell = row.AddCell()
	cell.Value = product.ShortNote.Zh
	cell = row.AddCell()
	cell.Value = product.ShortNote.En
	cell = row.AddCell()
	cell.Value = product.CurrencyCode
	cell = row.AddCell()
	cell.Value = "详细信息:"
	cell = row.AddCell()
	cell.Value = "身故赔偿:"
	cell = row.AddCell()
	cell.Value = scorez.N6
	cell = row.AddCell()
	cell.Value = scorez.N5
	cell = row.AddCell()
	cell.Value = scorez.N1
	cell = row.AddCell()
	cell.Value = scorez.N2
	cell = row.AddCell()
	cell.Value = scorez.N3
	cell = row.AddCell()
	cell.Value = scorez.N4
	cell = row.AddCell()
	cell.Value = "保证回报:"
	cell = row.AddCell()
	cell.Value = scorez.N18
	cell = row.AddCell()
	cell.Value = scorez.N14
	cell = row.AddCell()
	cell.Value = scorez.N7
	cell = row.AddCell()
	cell.Value = scorez.N8
	cell = row.AddCell()
	cell.Value = scorez.N9
	cell = row.AddCell()
	cell.Value = scorez.N10
	cell = row.AddCell()
	cell.Value = "预期回报:"
	cell = row.AddCell()
	cell.Value = scorez.N30
	cell = row.AddCell()
	cell.Value = scorez.N26
	cell = row.AddCell()
	cell.Value = scorez.N19
	cell = row.AddCell()
	cell.Value = scorez.N20
	cell = row.AddCell()
	cell.Value = scorez.N21
	cell = row.AddCell()
	cell.Value = scorez.N22
	cell = row.AddCell()
	cell.Value = "危机保障:"
	cell = row.AddCell()
	cell.Value = scorez.N40
	cell = row.AddCell()
	cell.Value = scorez.N35
	cell = row.AddCell()
	cell.Value = "癌症评分:"
	cell = row.AddCell()
	cell.Value = scorez.N41
	cell = row.AddCell()
	cell.Value = scorez.N36
	cell = row.AddCell()
	cell.Value = scorez.N49
	cell = row.AddCell()
	cell.Value = scorez.N48
	cell = row.AddCell()
	cell.Value = scorez.N50
	cell = row.AddCell()
	cell.Value = scorez.N51
	cell = row.AddCell()
	cell.Value = scorez.N63.Zh
	cell = row.AddCell()
	cell.Value = scorez.N63.En
	cell = row.AddCell()
	cell.Value = "心脏病评分:"
	cell = row.AddCell()
	cell.Value = scorez.N42
	cell = row.AddCell()
	cell.Value = scorez.N37
	cell = row.AddCell()
	cell.Value = scorez.N53
	cell = row.AddCell()
	cell.Value = scorez.N143
	cell = row.AddCell()
	cell.Value = scorez.N52
	cell = row.AddCell()
	cell.Value = scorez.N54
	cell = row.AddCell()
	cell.Value = scorez.N55
	cell = row.AddCell()
	cell.Value = scorez.N64.Zh
	cell = row.AddCell()
	cell.Value = scorez.N64.En
	cell = row.AddCell()
	cell.Value = "中风评分"
	cell = row.AddCell()
	cell.Value = scorez.N43
	cell = row.AddCell()
	cell.Value = scorez.N38
	cell = row.AddCell()
	cell.Value = scorez.N57
	cell = row.AddCell()
	cell.Value = scorez.N56
	cell = row.AddCell()
	cell.Value = scorez.N58
	cell = row.AddCell()
	cell.Value = scorez.N59
	cell = row.AddCell()
	cell.Value = scorez.N65.Zh
	cell = row.AddCell()
	cell.Value = scorez.N65.En
	cell = row.AddCell()
	cell.Value = "多次严重疾病评分"
	cell = row.AddCell()
	cell.Value = scorez.N44
	cell = row.AddCell()
	cell.Value = scorez.N39
	cell = row.AddCell()
	cell.Value = scorez.N147
	cell = row.AddCell()
	cell.Value = scorez.N148
	cell = row.AddCell()
	cell.Value = scorez.N150
	cell = row.AddCell()
	cell.Value = scorez.N149
	cell = row.AddCell()
	cell.Value = scorez.N151
	cell = row.AddCell()
	cell.Value = scorez.N152
	cell = row.AddCell()
	cell.Value = scorez.N153
	cell = row.AddCell()
	cell.Value = scorez.N66.Zh
	cell = row.AddCell()
	cell.Value = scorez.N66.En


	err = file.Save("储蓄型危疾.xlsx")
	if err != nil {
		panic(err)
	}
}

func stt2(){
	var s Productslice
	str := `{"products":[{"id":83,"name":{"en":"Personal Accident Insurance","zh-TW":"個人平安保險"},"short_note":{"en":"Based on Occupation Class 1 (indoor and non-hazadous nature)","zh-TW":"假設受保人從事辦公室及非危險性職務"},"payment_terms":{"en":null,"zh-TW":null},"issue_age":" - ","currency_code":"hkd","product_website":null,"scores":{"101":"1","102":"3","103":"3","104":"2","107":"7.778","108":"8.672","109":"8.654","110":"9.861","113":"1822","114":"1000000","115":"600","116":"40000","117":"548.8","118":"22.0","119":"2.3","130":"9","131":"9","132":"5.982","133":"4.982","134":{"en":"not covered","zh-TW":"不受保"},"135":"0","136":{"en":"$1000 per year","zh-TW":"全年上限$1,000"},"137":"100","138":{"en":"Yes","zh-TW":"包括"},"139":{"en":"No","zh-TW":"不包括"},"140":{"en":"Yes","zh-TW":"包括"},"141":{"en":"Yes","zh-TW":"包括"},"142":{"en":"No","zh-TW":"不包括"}},"insurance_company":{"id":17,"name":{"en":"QBE","zh-TW":"昆士蘭保險"},"logo_image_url":"assets/images/insurance_companies/qbe.png","product_category_ids":[3]},"product_category":{"id":3,"name":{"en":"Personal Accident","zh-TW":"個人意外"},"icon_name":"ios-walk"}},{"id":85,"name":{"en":"PAMultiple Personal Accident Plan","zh-TW":"樂在人生"},"short_note":{"en":"Based on Occupation Class 1 (indoor and non-hazadous nature)","zh-TW":"假設受保人從事辦公室及非危險性職務"},"payment_terms":{"en":null,"zh-TW":null},"issue_age":"18 - 70","currency_code":"hkd","product_website":null,"scores":{"101":"2","102":"6","103":"9","104":"10","107":"7.478","108":"8.333","109":"6.231","110":"3.146","113":"1896","114":"1000000","115":"200","116":"30000","117":"527.4","118":"15.8","119":"0.7","130":"4","131":"4","132":"9.981","133":"8.981","134":{"en":"$2000 per year overall","zh-TW":"全年上限$2,000"},"135":"80","136":{"en":"$2000 per year overall","zh-TW":"全年上限$2,000"},"137":"100","138":{"en":"Yes","zh-TW":"包括"},"139":{"en":"Yes","zh-TW":"包括"},"140":{"en":"Yes","zh-TW":"包括"},"141":{"en":"Yes","zh-TW":"包括"},"142":{"en":"Yes","zh-TW":"包括"}},"insurance_company":{"id":19,"name":{"en":"Zurich","zh-TW":"蘇黎世"},"logo_image_url":"assets/images/insurance_companies/zurich.png","product_category_ids":[3,4]},"product_category":{"id":3,"name":{"en":"Personal Accident","zh-TW":"個人意外"},"icon_name":"ios-walk"}},{"id":81,"name":{"en":"Family Personal Protector","zh-TW":"家庭個人意外保障計劃"},"short_note":{"en":"Based on Occupation Class 1 (indoor and non-hazadous nature)","zh-TW":"假設受保人從事辦公室及非危險性職務"},"payment_terms":{"en":null,"zh-TW":null},"issue_age":" - ","currency_code":"hkd","product_website":null,"scores":{"101":"3","102":"5","103":"14","104":"4","107":"7.280","108":"8.338","109":"4.150","110":"7.897","113":"1895","114":"1000000","115":"500","116":"20000","117":"527.7","118":"10.6","119":"1.8","130":"3","131":"5","132":"9.981","133":"8.731","134":{"en":"$500 per day upto $2500 per year","zh-TW":"每日上限$500, 全年上限$2,500"},"135":"100","136":{"en":"$150 per day upto $2500 per year","zh-TW":"每日上限$150, 全年上限$2,500"},"137":"75","138":{"en":"Yes","zh-TW":"包括"},"139":{"en":"Yes","zh-TW":"包括"},"140":{"en":"Yes","zh-TW":"包括"},"141":{"en":"Yes","zh-TW":"包括"},"142":{"en":"Yes","zh-TW":"包括"}},"insurance_company":{"id":16,"name":{"en":"MSIG","zh-TW":"MSIG"},"logo_image_url":"assets/images/insurance_companies/msig.png","product_category_ids":[3]},"product_category":{"id":3,"name":{"en":"Personal Accident","zh-TW":"個人意外"},"icon_name":"ios-walk"}},{"id":80,"name":{"en":"True Care Accident Protection Plan","zh-TW":"逸逸安心"},"short_note":{"en":"Based on Occupation Class 1 (indoor and non-hazadous nature)","zh-TW":"假設受保人從事辦公室及非危險性職務"},"payment_terms":{"en":null,"zh-TW":null},"issue_age":"0 - 65","currency_code":"hkd","product_website":null,"scores":{"101":"4","102":"11","103":"7","104":"1","107":"7.134","108":"5.267","109":"6.553","110":"9.970","113":"3000","114":"1000000","115":"1000","116":"50000","117":"333.3","118":"16.7","119":"2.3","130":"5","131":"6","132":"9.970","133":"7.970","134":{"en":"5 visits within provider network  per accident","zh-TW":"每次意外可向網絡指定醫療中心求診5次"},"135":"100","136":{"en":"5 visits within provider network  per accident","zh-TW":"每次意外可向網絡指定醫療中心求診5次"},"137":"100","138":{"en":"Yes","zh-TW":"包括"},"139":{"en":"Yes","zh-TW":"包括"},"140":{"en":"Yes","zh-TW":"包括"},"141":{"en":"Yes","zh-TW":"包括"},"142":{"en":"Yes","zh-TW":"包括"}},"insurance_company":{"id":10,"name":{"en":"MetLife","zh-TW":"大都會"},"logo_image_url":"assets/images/insurance_companies/metlife.png","product_category_ids":[1,2,3,4]},"product_category":{"id":3,"name":{"en":"Personal Accident","zh-TW":"個人意外"},"icon_name":"ios-walk"}},{"id":78,"name":{"en":"Personal Accident Insurance Plan","zh-TW":"人身意外"},"short_note":{"en":"Based on Occupation Class 1 (indoor and non-hazadous nature)","zh-TW":"假設受保人從事辦公室及非危險性職務"},"payment_terms":{"en":null,"zh-TW":null},"issue_age":"1 - 65","currency_code":"hkd","product_website":null,"scores":{"101":"5","102":"1","103":"1","104":"12","107":"6.989","108":"10.000","109":"9.984","110":"0.000","113":"1580","114":"1000000","115":"0","116":"40000","117":"632.9","118":"25.3","119":"0.0","130":"1","131":"12","132":"9.984","133":"0.000","134":{"en":"not covered","zh-TW":"不受保"},"135":"0","136":{"en":"not covered","zh-TW":"不受保"},"137":"0","138":{"en":"Yes","zh-TW":"包括"},"139":{"en":"Yes","zh-TW":"包括"},"140":{"en":"Yes","zh-TW":"包括"},"141":{"en":"Yes","zh-TW":"包括"},"142":{"en":"Yes","zh-TW":"包括"}},"insurance_company":{"id":15,"name":{"en":"Generali","zh-TW":"忠意保險"},"logo_image_url":"assets/images/insurance_companies/generali.png","product_category_ids":[3,4]},"product_category":{"id":3,"name":{"en":"Personal Accident","zh-TW":"個人意外"},"icon_name":"ios-walk"}},{"id":74,"name":{"en":"PA Insurance","zh-TW":"人意保"},"short_note":{"en":"Based on Occupation Class 1 (indoor and non-hazadous nature)","zh-TW":"假設受保人從事辦公室及非危險性職務"},"payment_terms":{"en":null,"zh-TW":null},"issue_age":"16 - 65","currency_code":"hkd","product_website":null,"scores":{"101":"6","102":"4","103":"4","104":"3","107":"6.758","108":"8.643","109":"8.625","110":"9.829","113":"1828","114":"1000000","115":"600","116":"40000","117":"547.0","118":"21.9","119":"2.3","130":"10","131":"13","132":"5.982","133":"0.000","134":{"en":"not covered","zh-TW":"不受保"},"135":"0","136":{"en":"not covered","zh-TW":"不受保"},"137":"0","138":{"en":"Yes","zh-TW":"包括"},"139":{"en":"No","zh-TW":"不包括"},"140":{"en":"Yes","zh-TW":"包括"},"141":{"en":"Yes","zh-TW":"包括"},"142":{"en":"No","zh-TW":"不包括"}},"insurance_company":{"id":14,"name":{"en":"China Pacific","zh-TW":"太平洋保險"},"logo_image_url":"assets/images/insurance_companies/cpic.png","product_category_ids":[3]},"product_category":{"id":3,"name":{"en":"Personal Accident","zh-TW":"個人意外"},"icon_name":"ios-walk"}},{"id":75,"name":{"en":"Personal Accident Insurance","zh-TW":"人身平安保險"},"short_note":{"en":"Based on Occupation Class 1 (indoor and non-hazadous nature)","zh-TW":"假設受保人從事辦公室及非危險性職務"},"payment_terms":{"en":null,"zh-TW":null},"issue_age":"16 - 65","currency_code":"hkd","product_website":null,"scores":{"101":"7","102":"2","103":"2","104":"14","107":"6.731","108":"9.405","109":"9.388","110":"0.000","113":"1680","114":"1000000","115":"0","116":"40000","117":"595.2","118":"23.8","119":"0.0","130":"12","131":"11","132":"1.983","133":"4.483","134":{"en":"$1000 per accident","zh-TW":"每次意外上限 $1,000 "},"135":"40","136":{"en":"$500 per accident","zh-TW":"每次意外上限 $500"},"137":"50","138":{"en":"No","zh-TW":"不包括"},"139":{"en":"No","zh-TW":"不包括"},"140":{"en":"Yes","zh-TW":"包括"},"141":{"en":"No","zh-TW":"不包括"},"142":{"en":"No","zh-TW":"不包括"}},"insurance_company":{"id":18,"name":{"en":"China Taiping","zh-TW":"中國太平"},"logo_image_url":"assets/images/insurance_companies/taiping.png","product_category_ids":[2,3]},"product_category":{"id":3,"name":{"en":"Personal Accident","zh-TW":"個人意外"},"icon_name":"ios-walk"}},{"id":73,"name":{"en":"Personal AccidentSafe","zh-TW":"個人意外至尊寶"},"short_note":{"en":"Based on Occupation Class 1 (indoor and non-hazadous nature)","zh-TW":"假設受保人從事辦公室及非危險性職務"},"payment_terms":{"en":null,"zh-TW":null},"issue_age":"16 - 65","currency_code":"hkd","product_website":null,"scores":{"101":"8","102":"10","103":"11","104":"6","107":"6.595","108":"5.623","109":"5.595","110":"6.378","113":"2810","114":"1000000","115":"600","116":"40000","117":"355.9","118":"14.2","119":"1.5","130":"11","131":"2","132":"5.972","133":"9.972","134":{"en":"$500 per visit upto $3,000 per year","zh-TW":"每次治療上限$500，全年$3,000上限"},"135":"100","136":{"en":"$200 per visit upto $3000 per year","zh-TW":"每次治療上限$200，全年$3,000上限"},"137":"100","138":{"en":"Yes","zh-TW":"包括"},"139":{"en":"Yes","zh-TW":"包括"},"140":{"en":"No","zh-TW":"不包括"},"141":{"en":"No","zh-TW":"不包括"},"142":{"en":"Yes","zh-TW":"包括"}},"insurance_company":{"id":13,"name":{"en":"Blue Cross","zh-TW":"藍十字保險"},"logo_image_url":"assets/images/insurance_companies/blue_cross.png","product_category_ids":[3,4]},"product_category":{"id":3,"name":{"en":"Personal Accident","zh-TW":"個人意外"},"icon_name":"ios-walk"}},{"id":82,"name":{"en":"Personal Accident","zh-TW":"安健保"},"short_note":{"en":"Based on Occupation Class 1 (indoor and non-hazadous nature)","zh-TW":"假設受保人從事辦公室及非危險性職務"},"payment_terms":{"en":null,"zh-TW":null},"issue_age":"3 - 65","currency_code":"hkd","product_website":null,"scores":{"101":"9","102":"8","103":"6","104":"5","107":"6.499","108":"6.605","109":"6.581","110":"7.501","113":"2392","114":"1000000","115":"600","116":"40000","117":"418.1","118":"16.7","119":"1.8","130":"7","131":"10","132":"7.976","133":"4.976","134":{"en":"not covered","zh-TW":"不受保"},"135":"0","136":{"en":"$200 per day upto $3000 per year","zh-TW":"每日上限$200, 全年上限$3,000"},"137":"100","138":{"en":"Yes","zh-TW":"包括"},"139":{"en":"Yes","zh-TW":"包括"},"140":{"en":"Yes","zh-TW":"包括"},"141":{"en":"Yes","zh-TW":"包括"},"142":{"en":"No","zh-TW":"不包括"}},"insurance_company":{"id":11,"name":{"en":"Prudential","zh-TW":"英國保誠"},"logo_image_url":"assets/images/insurance_companies/prudential.png","product_category_ids":[1,2,3,4]},"product_category":{"id":3,"name":{"en":"Personal Accident","zh-TW":"個人意外"},"icon_name":"ios-walk"}},{"id":87,"name":{"en":"PA Care Plus Select","zh-TW":"精選倍關心"},"short_note":{"en":"Based on Occupation Class 1 (indoor and non-hazadous nature)","zh-TW":"假設受保人從事辦公室及非危險性職務"},"payment_terms":{"en":null,"zh-TW":null},"issue_age":"18 - 70","currency_code":"hkd","product_website":null,"scores":{"101":"10","102":"12","103":"10","104":"13","107":"6.284","108":"4.883","109":"6.087","110":"0.000","113":"1618","114":"500000","115":"0","116":"25000","117":"309.0","118":"15.5","119":"0.0","130":"2","131":"1","132":"9.984","133":"9.984","134":{"en":"unlimited within provider network ","zh-TW":"不限次數向網絡指定醫療中心求診"},"135":"100","136":{"en":"unlimited within provider network ","zh-TW":"不限次數向網絡指定醫療中心求診"},"137":"100","138":{"en":"Yes","zh-TW":"包括"},"139":{"en":"Yes","zh-TW":"包括"},"140":{"en":"Yes","zh-TW":"包括"},"141":{"en":"Yes","zh-TW":"包括"},"142":{"en":"Yes","zh-TW":"包括"}},"insurance_company":{"id":1,"name":{"en":"AIA","zh-TW":"友邦"},"logo_image_url":"assets/images/insurance_companies/aia.png","product_category_ids":[1,2,3,4]},"product_category":{"id":3,"name":{"en":"Personal Accident","zh-TW":"個人意外"},"icon_name":"ios-walk"}},{"id":72,"name":{"en":"PAC Select 2","zh-TW":"自選人身意外保險2"},"short_note":{"en":"Based on Occupation Class 1 (indoor and non-hazadous nature)","zh-TW":"假設受保人從事辦公室及非危險性職務"},"payment_terms":{"en":null,"zh-TW":null},"issue_age":"18 - 70","currency_code":"hkd","product_website":null,"scores":{"101":"11","102":"13","103":"12","104":"8","107":"6.274","108":"4.620","109":"4.586","110":"5.229","113":"3420","114":"1000000","115":"600","116":"40000","117":"292.4","118":"11.7","119":"1.2","130":"6","131":"3","132":"9.966","133":"9.966","134":{"en":"Included in medical expense","zh-TW":"包括在意外受傷醫療保障額內"},"135":"100","136":{"en":"Included in medical expense","zh-TW":"包括在意外受傷醫療保障額內"},"137":"100","138":{"en":"Yes","zh-TW":"包括"},"139":{"en":"Yes","zh-TW":"包括"},"140":{"en":"Yes","zh-TW":"包括"},"141":{"en":"Yes","zh-TW":"包括"},"142":{"en":"Yes","zh-TW":"包括"}},"insurance_company":{"id":1,"name":{"en":"AIA","zh-TW":"友邦"},"logo_image_url":"assets/images/insurance_companies/aia.png","product_category_ids":[1,2,3,4]},"product_category":{"id":3,"name":{"en":"Personal Accident","zh-TW":"個人意外"},"icon_name":"ios-walk"}},{"id":84,"name":{"en":"Sun Care","zh-TW":"永關心意外保障計劃"},"short_note":{"en":"Based on Occupation Class 1 (indoor and non-hazadous nature)","zh-TW":"假設受保人從事辦公室及非危險性職務"},"payment_terms":{"en":null,"zh-TW":null},"issue_age":"18 - 65","currency_code":"hkd","product_website":null,"scores":{"101":"12","102":"9","103":"8","104":"9","107":"5.637","108":"6.371","109":"6.346","110":"4.295","113":"2480","114":"1000000","115":"357","116":"40000","117":"403.2","118":"16.1","119":"1.0","130":"15","131":"7","132":"-0.025","133":"6.975","134":{"en":"$1000 per year overall","zh-TW":"全年上限$1,000"},"135":"40","136":{"en":"$1000 per year overall","zh-TW":"全年上限$1,000"},"137":"100","138":{"en":"No","zh-TW":"不包括"},"139":{"en":"No","zh-TW":"不包括"},"140":{"en":"No","zh-TW":"不包括"},"141":{"en":"No","zh-TW":"不包括"},"142":{"en":"No","zh-TW":"不包括"}},"insurance_company":{"id":12,"name":{"en":"Sun Life","zh-TW":"永明"},"logo_image_url":"assets/images/insurance_companies/sunlife.png","product_category_ids":[1,2,3,4]},"product_category":{"id":3,"name":{"en":"Personal Accident","zh-TW":"個人意外"},"icon_name":"ios-walk"}},{"id":76,"name":{"en":"AccidentCare Plus Insurance","zh-TW":"綜合意外保險"},"short_note":{"en":"Based on Occupation Class 1 (indoor and non-hazadous nature)","zh-TW":"假設受保人從事辦公室及非危險性職務"},"payment_terms":{"en":null,"zh-TW":null},"issue_age":"18 - 65","currency_code":"hkd","product_website":null,"scores":{"101":"13","102":"7","103":"5","104":"11","107":"4.940","108":"7.900","109":"7.880","110":"2.123","113":"2000","114":"1000000","115":"143","116":"40000","117":"500.0","118":"20.0","119":"0.5","130":"14","131":"14","132":"-0.020","133":"0.000","134":{"en":"not covered","zh-TW":"不受保"},"135":"0","136":{"en":"not covered","zh-TW":"不受保"},"137":"0","138":{"en":"No","zh-TW":"不包括"},"139":{"en":"No","zh-TW":"不包括"},"140":{"en":"No","zh-TW":"不包括"},"141":{"en":"No","zh-TW":"不包括"},"142":{"en":"No","zh-TW":"不包括"}},"insurance_company":{"id":7,"name":{"en":"FWD","zh-TW":"富衛"},"logo_image_url":"assets/images/insurance_companies/fwd.png","product_category_ids":[1,2,3,4]},"product_category":{"id":3,"name":{"en":"Personal Accident","zh-TW":"個人意外"},"icon_name":"ios-walk"}},{"id":79,"name":{"en":"Personal Accident Plan","zh-TW":"個人意外計劃"},"short_note":{"en":"Based on Occupation Class 1 (indoor and non-hazadous nature)","zh-TW":"假設受保人從事辦公室及非危險性職務"},"payment_terms":{"en":null,"zh-TW":null},"issue_age":" - ","currency_code":"hkd","product_website":null,"scores":{"101":"14","102":"14","103":"13","104":"15","107":"4.023","108":"4.413","109":"4.378","110":"0.000","113":"3580","114":"1000000","115":"0","116":"40000","117":"279.3","118":"11.2","119":"0.0","130":"13","131":"8","132":"1.964","133":"5.964","134":{"en":"$300 per day upto medical expense limit","zh-TW":"每日上限$300，直至意外受傷醫療保障額上限"},"135":"60","136":{"en":"$600 per accident","zh-TW":"每次意外上限 $600"},"137":"60","138":{"en":"No","zh-TW":"不包括"},"139":{"en":"Yes","zh-TW":"包括"},"140":{"en":"No","zh-TW":"不包括"},"141":{"en":"No","zh-TW":"不包括"},"142":{"en":"No","zh-TW":"不包括"}},"insurance_company":{"id":8,"name":{"en":"Manulife","zh-TW":"宏利"},"logo_image_url":"assets/images/insurance_companies/manulife.png","product_category_ids":[1,2,3,4]},"product_category":{"id":3,"name":{"en":"Personal Accident","zh-TW":"個人意外"},"icon_name":"ios-walk"}},{"id":77,"name":{"en":"Attentive Care Plus Insurance","zh-TW":"綜合意外"},"short_note":{"en":"Based on Occupation Class 1 (indoor and non-hazadous nature)","zh-TW":"假設受保人從事辦公室及非危險性職務"},"payment_terms":{"en":null,"zh-TW":null},"issue_age":"0 - 60","currency_code":"hkd","product_website":null,"scores":{"101":"15","102":"15","103":"15","104":"7","107":"3.614","108":"3.285","109":"4.068","110":"6.199","113":"3848","114":"800000","115":"800","116":"40000","117":"207.9","118":"10.4","119":"1.5","130":"8","131":"15","132":"7.962","133":"0.000","134":{"en":"not covered","zh-TW":"不受保"},"135":"0","136":{"en":"not covered","zh-TW":"不受保"},"137":"0","138":{"en":"Yes","zh-TW":"包括"},"139":{"en":"Yes","zh-TW":"包括"},"140":{"en":"Yes","zh-TW":"包括"},"141":{"en":"Yes","zh-TW":"包括"},"142":{"en":"No","zh-TW":"不包括"}},"insurance_company":{"id":7,"name":{"en":"FWD","zh-TW":"富衛"},"logo_image_url":"assets/images/insurance_companies/fwd.png","product_category_ids":[1,2,3,4]},"product_category":{"id":3,"name":{"en":"Personal Accident","zh-TW":"個人意外"},"icon_name":"ios-walk"}}],"meta":{"next_page":null}}`
	json_err := json.Unmarshal([]byte(str), &s)
	if json_err != nil {
		fmt.Println(json_err)
	}
	//err := json.Unmarshal([]byte(str), &msgs)
	//SpecialPrint2(s.Products[1])


	for a:=0;a<15;a++{
		writeinmedicalscore(s.Products[a])
	}

}


