<template>
    <el-container class="designer-container">
      <el-header class="designer-header">
        <el-button type="primary" @click="previous"> 上一步 </el-button>
        <el-button type="primary" @click="saveSet"> 保存 </el-button>
      </el-header>
      <el-main class="designer-main">
        <div id="weboffice"
        style="height: 100%"
        :style="{width: isApplet ? '80%' : '100%'}"></div>
        <widgetsForm v-if="isApplet" @onInsert="onInsertHandler" />
      </el-main>
    </el-container>
  </template>
  
  <script setup>
  import { ElMessage } from 'element-plus'
  import * as WebOfficeSDK from '@/assets/js/web-office-sdk-solution-v1.1.24.es'
  import widgetsForm from '@/components/form-widget/index.vue'
  import { PcConfig, MobileConfig, objData } from './config'
  import { useFlowStore } from '@/store/workflow.js'
  import { generateRandom } from '@/utils/util'
  
  const flowState = useFlowStore()
  const instance = shallowRef()
  const app = shallowRef()
  const isApplet = ref(true)
  const contentControls = ref({})
  const asyncArrContorl = []
  
  const contractTemplateArr = computed(() => flowState.contractTemplateArr)
  
  // const { proxy } = getCurrentInstance()
  // const searchForm = ref({})
  
  // const loading = ref(false)
  // const query = ref({ page: 1, limit: 10 })
  // const totalCount = ref(0)
  // const tableData = ref()
  // const visible = ref(false)
  // const input1 = ref()
  // const input2 = ref()
  
  // // 加密
  // const str = window.btoa('1442370475882737665')
  // console.log('加密后', str)
  // // 解密
  // const str2 = window.atob(str)
  // console.log('解密后', str2)
  
  /**
   * 打开 webOffcie 时需要初始化的数据
   */
  const openWebOffice = async () => {
    instance.value = WebOfficeSDK.init({
      officeType: WebOfficeSDK.OfficeType.Writer, // Writer文档
      appId: 'SX20230713YRCIAQ',
      fileId: '1674639766377684994', // 文件 id
      token: '1442370475882737665', // 加密  1442370475882737665_1_2
      mount: '#weboffice', // 挂载
      mode: 'normal', // simple nomal
      wpsOptions: {
        isShowDocMap: false, // 是否开启目录功能，默认开启
        isBestScale: false, // 打开文档时，默认以最佳比例显示
        // isShowComment: false, // 是否显示评论相关入口
        // isShowBottomStatusBar: false, // 是否展示底部状态栏
        mobile: {
        // isOpenIntoEdit: false, // （Mobile）要有编辑权限，移动端打开时是否进入编辑
        // isShowHoverToolbars: false, // （Mobile）是否显示文字底部工具栏
          // isVoiceCommentEnabled: false, // （Mobile）是否允许插入语音评论
        },
      },
      commandBars: [
        {
          cmbId: 'HeaderLeft', // 组件 ID
          attributes: {
            visible: false, // 隐藏组件
            enable: false, // 禁用组件，组件显示但不响应点击事件
          },
        },
        {
          cmbId: 'HeaderRight', // 组件 ID
          attributes: {
            visible: false, // 隐藏组件
            enable: false, // 禁用组件，组件显示但不响应点击事件
          },
        },
        {
          cmbId: 'FloatQuickHelp', // 组件 ID
          attributes: {
            visible: false, // 隐藏组件
            enable: false, // 禁用组件，组件显示但不响应点击事件
          },
        },
        {
          cmbId: 'ContextMenuConvene', // 组件 ID
          attributes: {
            visible: false, // 隐藏组件
            enable: false, // 禁用组件，组件显示但不响应点击事件
          },
        },
        // {
        //   cmbId: 'TaskPane', // 组件 ID
        //   attributes: {
        //     visible: false, // 隐藏组件
        //     enable: false, // 禁用组件，组件显示但不响应点击事件
        //   },
        // },
  
      ],
      commonOptions: isApplet.value ? PcConfig : MobileConfig,
    })
  
    // 销毁 JSSDK 实例，关闭 WebOffice 文档的显示
    // await instance.destroy()()
    await instance.value.ready()
    app.value = instance.value.Application
  
    // 控制评论显示与否
    app.value.ActiveDocument.ActiveWindow.View.ShowComments = false
    // 是否在正文中显示评论
    // app.ActiveDocument.ActiveWindow.View.ShowComments = false
    // 隐藏评论相关功能（当前仅移动端有效）
    // await app.ActiveDocument.SetCommentEnabled({
    //   Enable: true,
    // })
    contentControls.value = await app.value.ActiveDocument.ContentControls
  
    // 选区对象
    // const selection = await app.ActiveDocument.ActiveWindow.Selection
  
    // Range 对象
    // const range = await selection.Range
    // console.log(range)
  
    // const Range = await app.value.ActiveDocument.Sections.Item(1).Headers.Range
    // console.log(Range)
    // Range.Text = '设置页眉'
  
    // 分页/连页模式切换
    if (!isApplet.value) {
      // 分页/连页模式切换
      await app.value.ActiveDocument.SwitchTypoMode(true)
    }
  
    // // 获取指定区域
    // const range = await app.ActiveDocument.Range(0, 100)
    // console.log(range)
  
    // 开启限制编辑
    // await app.ActiveDocument.Protect('test')
  
    // 停止限制编辑
    // await app.ActiveDocument.Unprotect('test')
  
    //
    // const controls = await app.CommandBars('FloatQuickHelp').Controls
    // const FloatQuickHelp = await controls.Add(1)
    // FloatQuickHelp.visible = false
    // console.log(FloatQuickHelp)
  
    // // 评论对象
    // const comments = await app.ActiveDocument.Comments
    // // 添加评论
    // await comments.Add({
    //   Range: {
    //     Start: 0,
    //     End: 9,
    //   },
    //   Text: 'WebOffice 评论',
    // })
  
    // const bookmarks = await app.ActiveDocument.Bookmarks
    // await bookmarks.Add({
    //   Name: 'WebOffice',
    //   Range: {
    //     Start: 1,
    //     End: 28,
    //   },
    // })
  
    // document.querySelector('.input1').addEventListener('input', async (e) => {
    //   // 内容控件对象
    //   const contentControls = await app.ActiveDocument.ContentControls
  
    //   // 获取第 1 个内容控件
    //   const contentControl = await contentControls.Item(1)
  
    //   // 获取第 1 个内容控件的范围
    //   const range = await contentControl.Range
  
    //   // 获取第 1 个内容控件的文本
    //   const text = range.Text
    //   console.log(text)
  
    //   // 设置第 1 个内容控件的文本
    //   range.Text = e.target.value
    // })
    // console.log(findResult)
  }
  
  /**
   * 插入
   */
  const onInsertHandler = async (index) => {
    try {
    // 在光标处插入内容控件
      const control = await contentControls.value.Add()
      console.log(control)
      const range = await control.Range
      console.log(range)
  
      // 返回第  1 个 ContentControl 对象
      // const contentControl = await contentControls.value.Item(1)
      // console.log(contentControl)
      // 获取第 1 个内容控件的范围
      // const range = await contentControl.Range
      // console.log(range)
  
      // // 获取第 1 个内容控件的文本
      // const text = range.Text
      // console.log(text)
  
      // const obj = objData[key]
      const obj = contractTemplateArr.value[index]
      console.log(obj)
  
      // range.HighlightColorIndex = 7
      // app.value.ActiveDocument.ActiveWindow.Selection.Range.HighlightColorIndex = 7
      // // 设置第 1 个内容控件的文本
      const ID = generateRandom()
      range.Text = `${obj.key}_${ID}`
      // range.Text = '设置第 1 个内容控件的文本'
  
      // const range2 = await control.Range
      // setTimeout(() => {
      //   range2.HighlightColorIndex = 3
      // const range = await app.ActiveDocument.Range(0, 100)
      // }, 100)
  
      const start = await range.Start
      console.log(start)
  
      const end = await range.End
      console.log(end)
  
      const range2 = await app.value.ActiveDocument.Range(start, end + String(ID).length)
      range2.HighlightColorIndex = 3
    } catch (error) {
      ElMessage({
        message: '请将光标定位到你要插入的位置！！',
        type: 'warning',
      })
    }
    // 获取第 1 个内容控件占位符文本
    // const placeholderText = await contentControl.PlaceholderText
    // console.log(placeholderText)
  }
  
  async function printFiles() {
    await Promise.all(asyncArrContorl).then(async (range) => {
      // console.log(res)
      const text = await range.Text
      console.log(text)
    })
    // await Promise.all(asyncArrContorl.map(async (range) => {
    //   const text = await range.Text
    //   console.log(range)
    // }))
  }
  
  const getPlaceholderText = async (i) => {
    // const contentControl = await contentControls.value.Item(i)
    // const placeholderText = await contentControl.Placeho
    // console.log(placeholderText)
    // 获取第 1 个内容控件
    try {
      const contentControl = await contentControls.value.Item(i)
  
      // 获取选中文本
      const range = await contentControl.Range
  
      // // 获取第 1 个内容控件的文本
      const text = await range.Text
      // console.log(text)
  
      // // 查找对应的 key
      if (text) {
        console.log(123)
        const key = objData.find((item) => item.text === text.split('_')[0])
        console.log(key)
        range.Text = key.value
      }
    } catch (error) {
      return undefined
    }
  }
  
  /**
   * 保存替换
   */
  const saveSet = async () => {
    // 内容控件数量
    const count = await contentControls.value.Count
    console.log(count) // 5
    // 获取选中区域
    // const DocumentRange = await app.value.ActiveDocument.GetDocumentRange()
    // console.log(DocumentRange.Count)
    // 区域中单元格的数量
    // const count = await DocumentRange.Count
    // console.log(count)
    // 获取第 1 个内容控件
    // const contentControl = await contentControls.value.Item(1)
  
    // // 获取选中文本
    // const range = await contentControl.Range
  
    // // 获取第 1 个内容控件的文本
    // const text = await range.Text
    // console.log(text)
    // 获取第 1 个内容控件占位符文本
    //  const placeholderText = await contentControl.PlaceholderText
    // console.log(placeholderText)
  
    // // 设置第 1 个内容控件的文本
    // range.Text = e.target.value
  
    // const contentControl = await contentControls.value.Item(2)
    // const placeholderText = await contentControl.Placeho
  
    // console.log(placeholderText)
  
    if (count > 0) {
      for (let i = 0; i < count; i++) {
        // eslint-disable-next-line no-await-in-loop
        await getPlaceholderText(i)
        // try {
        //   const contentControl = await contentControls.value.Item(i)
  
        //   // 获取选中文本
        //   const range = await contentControl.Range
        //   const text = await range.Text
        //   console.log(text)
        //   const key = objData.find((item) => item.text === text.split('_')[0])
        //   range.Text = key.value
        // } catch (error) {
        //   //
        // }
      }
      printFiles()
    }
  }
  
  /**
   * 判断当前是否为 web-view环境
   */
  const init = () => {
    function ready() {
      // console.log(window.__wxjs_environment === 'miniprogram') // true
      // eslint-disable-next-line no-underscore-dangle
      // alert(window.__wxjs_environment === 'miniprogram')
      // eslint-disable-next-line no-underscore-dangle
      if (window.__wxjs_environment === 'miniprogram') {
        isApplet.value = false
      }
    }
    // eslint-disable-next-line no-undef
    if (!window.WeixinJSBridge || !WeixinJSBridge.invoke) {
      document.addEventListener('WeixinJSBridgeReady', ready, false)
      console.log(123)
    } else {
      ready()
    }
  }
  init()
  
  onMounted(() => {
    openWebOffice()
  })
  </script>
  
  <style lang="scss" scoped>
  @import "./index.scss";
  </style>
  