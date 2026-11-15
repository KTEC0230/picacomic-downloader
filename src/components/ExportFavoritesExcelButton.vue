<script setup lang="ts">
import { onMounted, ref } from 'vue'
import { MessageReactive, useMessage } from 'naive-ui'
import { save } from '@tauri-apps/plugin-dialog'
import { commands, events, GetFavoriteSort } from '../bindings.ts'

const props = defineProps<{ sort: GetFavoriteSort }>()

const message = useMessage()

const exporting = ref(false)
let loadingMessage: MessageReactive | undefined

onMounted(async () => {
  await events.exportFavoritesExcelEvent.listen(({ payload }) => {
    if (payload.event === 'GettingFavorites') {
      const { current, total } = payload.data
      if (loadingMessage === undefined) {
        loadingMessage = message.loading('正在获取收藏夹', { duration: 0 })
      }
      loadingMessage.content = `正在获取收藏夹(${current}/${total}页)`
    } else if (payload.event === 'End' && loadingMessage !== undefined) {
      const { comicCount, savePath } = payload.data
      loadingMessage.type = 'success'
      loadingMessage.content = `成功导出${comicCount}部漫画到 ${savePath}`
      const msg = loadingMessage
      setTimeout(() => {
        if (loadingMessage === msg) {
          loadingMessage?.destroy()
          loadingMessage = undefined
        }
      }, 5000)
    }
  })
})

async function handleExportClick() {
  if (exporting.value) {
    return
  }

  // 默认文件名带时间戳，如`哔咔收藏夹_20261115_173000.xlsx`
  const now = new Date()
  const pad = (n: number) => String(n).padStart(2, '0')
  const timestamp = `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`

  // 弹出保存对话框，让用户自行选择输出路径与文件名
  const savePath = await save({
    title: '选择收藏夹Excel的保存位置',
    defaultPath: `哔咔收藏夹_${timestamp}.xlsx`,
    filters: [{ name: 'Excel 工作簿', extensions: ['xlsx'] }],
  })
  if (savePath === null) {
    return
  }

  exporting.value = true
  loadingMessage = message.loading('正在获取收藏夹', { duration: 0 })
  const result = await commands.exportFavoritesExcel(props.sort, savePath)
  if (result.status === 'error') {
    console.error(result.error)
    loadingMessage.destroy()
    loadingMessage = undefined
    message.error(`导出收藏夹Excel失败: ${result.error.err_title}`, { duration: 5000 })
  }
  exporting.value = false
}
</script>

<template>
  <n-button size="small" :disabled="exporting" @click="handleExportClick">导出收藏夹Excel</n-button>
</template>
